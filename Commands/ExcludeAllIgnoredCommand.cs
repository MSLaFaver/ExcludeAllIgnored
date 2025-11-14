using EnvDTE;
using Microsoft.Build.Evaluation;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcludeAllIgnored
{
	[Command(PackageIds.ExcludeAllIgnoredCommand)]
	internal sealed class ExcludeAllIgnoredCommand : BaseCommand<ExcludeAllIgnoredCommand>
	{
		protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
		{

			// Get DTE and solution (must be on UI thread)
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
			var dte = (DTE)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
			var solution = dte.Solution;

			// Collect all files currently included in projects
			var allFiles = new List<string>();
			foreach (EnvDTE.Project proj in solution.Projects)
			{
				try
				{
					// Inline traversal of project items (previously in CollectItems)
					if (proj.ProjectItems != null)
					{
						var stack = new System.Collections.Generic.Stack<EnvDTE.ProjectItems>();
						stack.Push(proj.ProjectItems);
						while (stack.Count > 0)
						{
							var items = stack.Pop();
							if (items == null) continue;
							foreach (EnvDTE.ProjectItem item in items)
							{
								try
								{
									for (short i = 1; i <= item.FileCount; i++)
									{
										try
										{
											var name = item.FileNames[i];
											if (!string.IsNullOrEmpty(name) && File.Exists(name)) allFiles.Add(Path.GetFullPath(name));
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
				catch
				{
					// ignore individual project errors
				}
			}

			if (!allFiles.Any())
			{
				await VS.MessageBox.ShowAsync("No files found in open solution projects.");
				return;
			}

			// Group files by git repo root (walk up looking for .git)
			var filesByRepo = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
			foreach (var file in allFiles)
			{
				var repo = FindGitRoot(Path.GetDirectoryName(file));
				if (repo == null) continue; // not tracked by any repo
				if (!filesByRepo.TryGetValue(repo, out var list))
				{
					list = [];
					filesByRepo[repo] = list;
				}
				list.Add(file);
			}

			if (!filesByRepo.Any())
			{
				await VS.MessageBox.ShowAsync("No files under a git repository were found.");
				return;
			}

			var ignoredFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

			// For each repo, run `git check-ignore --stdin -z` to find ignored files among the list
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
				catch (Exception ex)
				{
					// ignore git failures for a repo
					Debug.WriteLine(ex);
				}
			}

			if (!ignoredFiles.Any())
			{
				await VS.MessageBox.ShowAsync("No ignored files found among project items.");
				return;
			}

			// Group ignored files by project file path and remove their corresponding MSBuild items
			var projects = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
			foreach (EnvDTE.Project proj in solution.Projects)
			{
				try
				{
					var projPath = proj.FullName;
					if (string.IsNullOrEmpty(projPath) || !File.Exists(projPath)) continue;
					var projDir = Path.GetDirectoryName(projPath);
					foreach (var f in ignoredFiles)
					{
						if (Path.GetFullPath(f).StartsWith(projDir + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || string.Equals(Path.GetFullPath(f), Path.GetFullPath(projDir), StringComparison.OrdinalIgnoreCase))
						{
							if (!projects.TryGetValue(projPath, out var list)) { list = new List<string>(); projects[projPath] = list; }
							list.Add(f);
						}
					}
				}
				catch { }
			}

			var changedProjects = new List<string>();

			foreach (var kv in projects)
			{
				var projPath = kv.Key;
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
						// Try to match items whose evaluated include resolves to this file
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
							msproj.RemoveItem(item);
							anyRemoved = true;
						}
					}

					if (anyRemoved)
					{
						msproj.Save();
						changedProjects.Add(projPath);
					}
				}
				catch (Exception ex)
				{
					Debug.WriteLine(ex);
				}
			}

			var msg = new StringBuilder();
			msg.AppendLine($"Ignored files found: {ignoredFiles.Count}");
			msg.AppendLine($"Projects modified: {changedProjects.Count}");
			await VS.MessageBox.ShowAsync(msg.ToString());
		}

		// CollectItems has been inlined at call sites to avoid a separate method.

		private static string FindGitRoot(string startDir)
		{
			var dir = startDir;
			while (!string.IsNullOrEmpty(dir))
			{
				if (Directory.Exists(Path.Combine(dir, ".git"))) return dir;
				var parent = Directory.GetParent(dir);
				dir = parent?.FullName;
			}
			return null;
		}

		private static List<string> RunGitCheckIgnore(string workingDirectory, List<string> relativePaths)
		{
			var result = new List<string>();
			if (relativePaths == null || relativePaths.Count == 0) return result;

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
				// write all paths separated by NUL
				var stdin = p.StandardInput.BaseStream;
				var enc = Encoding.UTF8;
				foreach (var r in relativePaths)
				{
					var bytes = enc.GetBytes(r + "\0");
					stdin.Write(bytes, 0, bytes.Length);
				}
				stdin.Flush();
				p.StandardInput.Close();

				// read all output into memory
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

			return result;
		}

		private static string GetRelativePath(string basePath, string path)
		{
			try
			{
				var baseUri = new Uri(AppendDirectorySeparatorChar(Path.GetFullPath(basePath)));
				var pathUri = new Uri(Path.GetFullPath(path));
				var rel = baseUri.MakeRelativeUri(pathUri).ToString();
				rel = Uri.UnescapeDataString(rel).Replace('/', Path.DirectorySeparatorChar);
				return rel;
			}
			catch
			{
				// fallback
				return path;
			}
		}

		private static string AppendDirectorySeparatorChar(string path)
		{
			if (!path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
				return path + Path.DirectorySeparatorChar;
			return path;
		}
	}
}
