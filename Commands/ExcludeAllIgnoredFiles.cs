using EnvDTE;
using Microsoft.Build.Evaluation;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcludeAllIgnored
{
	[Command(PackageIds.ExcludeAllIgnoredFiles)]
	internal sealed class ExcludeAllIgnoredFiles : BaseCommand<ExcludeAllIgnoredFiles>
	{
		protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
		{
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
			var dte = (DTE)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
			var solution = dte.Solution;

			var allFiles = new List<string>();
			var msg = new StringBuilder();
			foreach (EnvDTE.Project proj in solution.Projects)
			{
				try
				{
					var beforeCount = allFiles.Count;
					if (proj.ProjectItems != null)
					{
						var stack = new Stack<ProjectItems>();
						stack.Push(proj.ProjectItems);
						while (stack.Count > 0)
						{
							var items = stack.Pop();
							if (items != null)
							{
								foreach (EnvDTE.ProjectItem item in items)
								{
									try
									{
										for (short i = 1; i <= item.FileCount; i++)
										{
											try
											{
												var name = item.FileNames[i];
												if (!string.IsNullOrEmpty(name) && File.Exists(name))
												{
													allFiles.Add(Path.GetFullPath(name));
												}
											}
											catch { }
										}

										if (item.ProjectItems != null && item.ProjectItems.Count > 0)
										{
											stack.Push(item.ProjectItems);
										}
									}
									catch { }
								}
							}
						}
					}
					var added = allFiles.Count - beforeCount;
				}
				catch { }
			}

			if (!allFiles.Any())
			{
				msg.AppendLine($"Warning: No valid projects with files found.");
				await VS.MessageBox.ShowAsync(msg.ToString());
			}
			else
			{
				var filesByRepo = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
				foreach (var file in allFiles)
				{
					var repo = FindGitRoot(Path.GetDirectoryName(file));
					if (repo != null)
					{
						if (!filesByRepo.TryGetValue(repo, out var list))
						{
							list = [];
							filesByRepo[repo] = list;
						}
						list.Add(file);
					}
				}

				if (!filesByRepo.Any())
				{
					msg.AppendLine("No files found specified as \"ignore\" in Git repo.");
					await VS.MessageBox.ShowAsync(msg.ToString());
				}
				else
				{
					var ignoredFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

					foreach (var kv in filesByRepo)
					{
						var repoRoot = kv.Key;
						var files = kv.Value;

						try
						{
							var relPaths = files.Select(f => GetRelativePath(repoRoot, f).Replace(Path.DirectorySeparatorChar, '/')).ToList();
							var matched = RunGitIgnoreCheck(repoRoot, relPaths);
							foreach (var r in matched)
							{
								var abs = Path.GetFullPath(Path.Combine(repoRoot, r));
								ignoredFiles.Add(abs);
							}
						}
						catch { }
					}

					if (!ignoredFiles.Any())
					{
						msg.AppendLine("The tool has run and no ignored files were found as included.");
						await VS.MessageBox.ShowAsync(msg.ToString());
					}
					else
					{
						var projects = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
						foreach (EnvDTE.Project proj in solution.Projects)
						{
							try
							{
								var projPath = proj.FullName;
								if (!string.IsNullOrEmpty(projPath) && File.Exists(projPath))
								{
									var projDir = Path.GetDirectoryName(projPath);
									foreach (var f in ignoredFiles)
									{
										if (Path.GetFullPath(f).StartsWith(projDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || string.Equals(Path.GetFullPath(f), Path.GetFullPath(projDir), StringComparison.OrdinalIgnoreCase))
										{
											if (!projects.TryGetValue(projPath, out var list))
											{
												list = [];
												projects[projPath] = list;
											}
											list.Add(f);
										}
									}
								}
							}
							catch { }
						}

						var changedProjects = new List<string>();

						foreach (var kv in projects)
						{
							var projPath = kv.Key;
							msg.AppendLine($"Excluding the following files from {projPath}:");

							var toRemove = kv.Value.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

							try
							{
								var pc = ProjectCollection.GlobalProjectCollection;
								var msproj = pc.LoadProject(projPath);
								var projDir = Path.GetDirectoryName(projPath);

								bool anyRemoved = false;
								foreach (var file in toRemove)
								{
									var rel = GetRelativePath(projDir, file).Replace(Path.DirectorySeparatorChar, '\\');
									var matches = msproj.Items.Where(i =>
									{
										try
										{
											var evaluated = i.EvaluatedInclude ?? string.Empty;
											var full = Path.GetFullPath(Path.Combine(projDir, evaluated));
											return string.Equals(full, Path.GetFullPath(file), StringComparison.OrdinalIgnoreCase);
										}
										catch { return false; }
									}).ToList();

									foreach (var item in matches)
									{
										var evaluated = item.EvaluatedInclude ?? "";
										var fullPath = Path.GetFullPath(Path.Combine(projDir, evaluated));
										msproj.RemoveItem(item);
										msg.AppendLine($"\t{fullPath}");
										anyRemoved = true;
									}
								}

								if (anyRemoved)
								{
									msproj.Save();
									changedProjects.Add(projPath);
								}
							}
							catch { }

							ProjectCollection.GlobalProjectCollection.UnloadAllProjects();
						}

						msg.AppendLine("\nYou may now reload the project to apply the changes.");
						await VS.MessageBox.ShowAsync(msg.ToString());
					}
				}
			}
		}

		private static string FindGitRoot(string startDir)
		{
			var dir = startDir;
			string ret = null;
			while (!string.IsNullOrEmpty(dir))
			{
				if (Directory.Exists(Path.Combine(dir, ".git")))
				{
					ret = dir;
					break;
				}
				var parent = Directory.GetParent(dir);
				dir = parent?.FullName;
			}
			return ret;
		}

		private static List<string> RunGitIgnoreCheck(string workingDirectory, List<string> relativePaths)
		{
			var result = new List<string>();
			if (!(relativePaths == null || relativePaths.Count == 0))
			{
				var gi = Path.Combine(workingDirectory, ".gitignore");
				if (File.Exists(gi))
				{
					var patterns = ReadGitIgnorePatterns(gi);

					foreach (var rp in relativePaths)
					{
						var rel = rp.Replace('\\', '/');
						bool ignored = false;
						foreach (var pat in patterns)
						{
							if (string.IsNullOrEmpty(pat))
							{
								continue;
							}
							var neg = pat.StartsWith("!");
							var p = neg ? pat.Substring(1) : pat;
							if (MatchesGitIgnorePattern(p, rel))
							{
								if (neg)
								{
									ignored = false;
								}
								else
								{
									ignored = true;
								}
							}
						}
						if (ignored)
						{
							result.Add(rel);
						}
					}
				}
			}
			return result;
		}

		private static List<string> ReadGitIgnorePatterns(string gitignorePath)
		{
			var lines = new List<string>();
			foreach (var raw in File.ReadAllLines(gitignorePath))
			{
				var l = raw.Trim();
				if (!(string.IsNullOrEmpty(l) || l.StartsWith("#")))
				{
					lines.Add(l);
				}
			}
			return lines;
		}

		private static bool MatchesGitIgnorePattern(string pattern, string relPath)
		{
			var pat = pattern.Replace('\\', '/');
			var path = relPath.Replace('\\', '/');

			bool matched = false;

			bool directoryPattern = pat.EndsWith("/");
			if (directoryPattern)
			{
				pat = pat.Substring(0, pat.Length - 1);
			}

			bool anchored = pat.StartsWith("/");
			if (anchored)
			{
				pat = pat.Substring(1);
			}

			if (pat == "")
			{
				matched = false;
			}
			else
			{
				var patRegex = ConvertGlobToRegex(pat);

				if (!matched && anchored)
				{
					var rx = new Regex("^" + patRegex + "($|/.*)", RegexOptions.IgnoreCase);
					if (rx.IsMatch(path))
					{
						matched = true;
					}
				}

				if (!matched && pat.Contains("/"))
				{
					var rx = new Regex("(^|.*/)" + patRegex + "($|/.*)", RegexOptions.IgnoreCase);
					if (rx.IsMatch(path))
					{
						matched = true;
					}
				}

				if (!matched)
				{
					var fileName = Path.GetFileName(path);
					var rxf = new Regex("^" + patRegex + "$", RegexOptions.IgnoreCase);
					if (rxf.IsMatch(fileName))
					{
						matched = true;
					}
					else
					{
						var segments = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
						foreach (var seg in segments)
						{
							if (rxf.IsMatch(seg))
							{
								matched = true;
								break;
							}
						}
					}
				}
			}

			return matched;
		}

		private static string ConvertGlobToRegex(string pattern)
		{
			var esc = Regex.Escape(pattern);
			esc = esc.Replace("\\*\\*", "###DS###");
			esc = esc.Replace("\\*", "###S###");
			esc = esc.Replace("\\?", "###Q###");
			esc = esc.Replace("###DS###", ".*");
			esc = esc.Replace("###S###", "[^/]*");
			esc = esc.Replace("###Q###", "[^/]");
			return esc;
		}


		private static string GetRelativePath(string basePath, string path)
		{
			string ret = path;
			try
			{
				var baseUri = new Uri(AppendDirectorySeparatorChar(Path.GetFullPath(basePath)));
				var pathUri = new Uri(Path.GetFullPath(path));
				var rel = baseUri.MakeRelativeUri(pathUri).ToString();
				rel = Uri.UnescapeDataString(rel).Replace('/', Path.DirectorySeparatorChar);
				ret = rel;
			}
			catch
			{
				ret = path;
			}
			return ret;
		}

		private static string AppendDirectorySeparatorChar(string path)
		{
			string ret = path;
			if (!path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
			{
				ret = path + Path.DirectorySeparatorChar;
			}
			return ret;
		}
	}
}