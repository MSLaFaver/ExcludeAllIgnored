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
						var stack = new System.Collections.Generic.Stack<EnvDTE.ProjectItems>();
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
							list = new List<string>();
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
							var matched = RunGitCheckIgnore(repoRoot, relPaths);
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
												list = new List<string>(); projects[projPath] = list;
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
											return IncludeMatchesFile(projDir, evaluated, file);
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

									pc.UnloadAllProjects();
								}

								if (anyRemoved)
								{
									msproj.Save();
									changedProjects.Add(projPath);
								}
							}
							catch { }
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

		private static List<string> RunGitCheckIgnore(string workingDirectory, List<string> relativePaths)
		{
			var result = new List<string>();
			if (!(relativePaths == null || relativePaths.Count == 0))
			{
				var psi = new System.Diagnostics.ProcessStartInfo("git")
				{
					Arguments = "check-ignore --stdin -z",
					RedirectStandardInput = true,
					RedirectStandardOutput = true,
					RedirectStandardError = true,
					UseShellExecute = false,
					CreateNoWindow = true,
					WorkingDirectory = workingDirectory,
				};

				using (var p = System.Diagnostics.Process.Start(psi))
				{
					var stdin = p.StandardInput.BaseStream;
					var enc = Encoding.UTF8;
					foreach (var r in relativePaths)
					{
						var bytes = enc.GetBytes(r + "\0");
						stdin.Write(bytes, 0, bytes.Length);
					}
					stdin.Flush();
					p.StandardInput.Close();

					using (var ms = new MemoryStream())
					{
						p.StandardOutput.BaseStream.CopyTo(ms);
						var outBytes = ms.ToArray();
						if (outBytes != null && outBytes.Length > 0)
						{
							var outStr = Encoding.UTF8.GetString(outBytes);
							var parts = outStr.Split(new char[] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
							result.AddRange(parts);
						}
					}

					p.WaitForExit(5000);
				}
			}

			return result;
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

		private static bool IncludeMatchesFile(string projDir, string evaluatedInclude, string filePath)
		{
            bool result;
            try
			{
				if (string.IsNullOrEmpty(evaluatedInclude))
				{
					var full = Path.GetFullPath(Path.Combine(projDir, evaluatedInclude));
					result = string.Equals(full, Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase);
				}
				else if (evaluatedInclude.IndexOfAny(['*', '?']) < 0 && !evaluatedInclude.Contains("**"))
				{
					var full = Path.GetFullPath(Path.Combine(projDir, evaluatedInclude));
					result = string.Equals(full, Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase);
				}
				else
				{
					var patternPath = Path.Combine(projDir, evaluatedInclude).Replace('/', Path.DirectorySeparatorChar);
					var sb = new StringBuilder();
					sb.Append('^');
					for (int i = 0; i < patternPath.Length; i++)
					{
						var c = patternPath[i];
						if (c == '*')
						{
							if (i + 1 < patternPath.Length && patternPath[i + 1] == '*')
							{
								sb.Append(".*");
								i++;
							}
							else
							{
								sb.Append("[^\\\\]*");
							}
						}
						else if (c == '?')
						{
							sb.Append("[^\\\\]");
						}
						else
						{
							sb.Append(Regex.Escape(c.ToString()));
						}
					}
					sb.Append('$');

					var regex = new Regex(sb.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
					var fullFile = Path.GetFullPath(filePath);
					result = regex.IsMatch(fullFile);
				}
			}
			catch
			{
				result = false;
			}
			return result;
		}
	}
}