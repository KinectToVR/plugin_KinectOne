using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage;
using Amethyst.Plugins.Contract;
using Microsoft.UI.Xaml.Controls;
using RestSharp;

namespace plugin_KinectOne;

internal class SetupData : ICoreSetupData
{
    public object PluginIcon => new BitmapIcon
    {
        UriSource = new Uri(Path.Join(Directory.GetParent(
                Assembly.GetExecutingAssembly().Location)!.FullName,
            "Assets", "Resources", "icon.png"))
    };

    public string GroupName => "kinect";
    public Type PluginType => typeof(ITrackingDevice);
}

internal class RuntimeInstaller : IDependencyInstaller
{
    private const string WixDownloadUrl =
        "https://github.com/wixtoolset/wix3/releases/download/wix3112rtm/wix311-binaries.zip";

    private const string RuntimeDownloadUrl =
        "https://download.microsoft.com/download/A/7/4/A74239EB-22C2-45A1-996C-2F8E564B28ED/KinectRuntime-v2.0_1409-Setup.exe";

    private string TemporaryFolderName { get; } = Guid.NewGuid().ToString().ToUpper();

    public Task<bool> InstallTools(IProgress<InstallationProgress> progress)
    {
        return Task.FromResult(false); // Not supported (yet?)
    }

    public bool IsInstalled
    {
        get
        {
            try
            {
                // Well, this is pretty much all we need for the plugin to be loaded
                return File.Exists(
                    @"C:\Windows\Microsoft.NET\assembly\GAC_64\Microsoft.Kinect\v4.0_2.0.0.0__31bf3856ad364e35\Microsoft.Kinect.dll");
            }
            catch (Exception)
            {
                // Access denied?
                return false;
            }
        }
    }

    public string InstallerEula
    {
        get
        {
            try
            {
                return File.ReadAllText(Path.Join(
                    Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.FullName,
                    "Assets", "Resources", "eula.md"));
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
    }

    public bool ProvidesTools => false;
    public bool ToolsInstalled => false;
    public IDependencyInstaller.ILocalizationHost Host { get; set; }

    public async Task<bool> Install(IProgress<InstallationProgress> progress)
    {
        return
            // Download and unpack WiX
            await SetupWix("WiXToolset", progress) &&

            // Download, unpack, and install the runtime
            await SetupRuntime("WiXToolset", progress);
    }

    private async Task<StorageFolder> GetTempDirectory()
    {
        return await ApplicationData.Current.TemporaryFolder.CreateFolderAsync(
            TemporaryFolderName, CreationCollisionOption.OpenIfExists);
    }

    private async Task<bool> SetupWix(string outputFolder, IProgress<InstallationProgress> progress)
    {
        try
        {
            using var client = new RestClient();
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Downloading/WiX") ??
                             "Downloading WiX Toolset"
            });

            // Create a stream reader using the received Installer Uri
            await using var stream =
                await client.ExecuteDownloadStreamAsync(WixDownloadUrl, new RestRequest());

            // Replace or create our installer file
            var installerFile = await (await GetTempDirectory()).CreateFileAsync(
                "wix-binaries.zip", CreationCollisionOption.ReplaceExisting);

            // Create an output stream and push all the available data to it
            await using var fsInstallerFile = await installerFile.OpenStreamForWriteAsync();
            await stream.CopyToWithProgressAsync(fsInstallerFile, innerProgress =>
            {
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = false, OverallProgress = innerProgress / 34670186.0,
                    StageTitle = Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Downloading/WiX") ??
                                 "Downloading WiX Toolset"
                });
            }); // The runtime will do the rest for us

            // Close the file to unlock it
            fsInstallerFile.Close();

            var sourceZip = Path.GetFullPath(Path.Combine((await GetTempDirectory()).Path, "wix-binaries.zip"));
            var tempDirectory = Path.GetFullPath(Path.Combine((await GetTempDirectory()).Path, outputFolder));

            if (File.Exists(sourceZip))
            {
                if (!Directory.Exists(tempDirectory))
                    Directory.CreateDirectory(tempDirectory);

                try
                {
                    // Extract the toolset
                    ZipFile.ExtractToDirectory(sourceZip, tempDirectory, true);
                }
                catch (Exception e)
                {
                    progress.Report(new InstallationProgress
                    {
                        IsIndeterminate = true,
                        StageTitle =
                            (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Exceptions/WiX/Extraction") ??
                             "Toolset extraction failed! Exception: {0}").Replace("{0}", e.Message)
                    });

                    return false;
                }

                return true;
            }
        }
        catch (Exception e)
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Exceptions/WiX/Installation") ??
                              "Toolset installation failed! Exception: {0}").Replace("{0}", e.Message)
            });
            return false;
        }

        return false;
    }

    private async Task<bool> SetupRuntime(string wixFolder, IProgress<InstallationProgress> progress)
    {
        try
        {
            using var client = new RestClient();
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Downloading/Runtime") ??
                             "Downloading Kinect for Xbox One Runtime..."
            });

            // Create a stream reader using the received Installer Uri
            await using var stream =
                await client.ExecuteDownloadStreamAsync(RuntimeDownloadUrl, new RestRequest());

            // Replace or create our installer file
            var installerFile = await (await GetTempDirectory()).CreateFileAsync(
                "kinect-setup.exe", CreationCollisionOption.ReplaceExisting);

            // Create an output stream and push all the available data to it
            await using var fsInstallerFile = await installerFile.OpenStreamForWriteAsync();
            await stream.CopyToWithProgressAsync(fsInstallerFile, innerProgress =>
            {
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = false,
                    OverallProgress = innerProgress / 93314296.0,
                    StageTitle = Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Downloading/Runtime") ??
                                 "Downloading Kinect for Xbox One Runtime..."
                });
            }); // The runtime will do the rest for us

            // Close the file to unlock it
            fsInstallerFile.Close();

            return
                // Extract all runtime files for the installation
                await ExtractFiles(Path.GetFullPath(Path.Combine((await GetTempDirectory()).Path, wixFolder)),
                    installerFile.Path, Path.Join((await GetTempDirectory()).Path, "KinectRuntime"), progress) &&

                // Install the files using msi installers
                InstallFiles(Directory.GetFiles(Path.Join((await GetTempDirectory()).Path,
                    "KinectRuntime", "AttachedContainer"), "*.msi"), progress);
        }
        catch (Exception e)
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle =
                    (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Exceptions/Runtime/Installation") ??
                     "Runtime installation failed! Exception: {0}").Replace("{0}", e.Message)
            });
            return false;
        }
    }

    private async Task<bool> ExtractFiles(string wixPath, string sourceFile, string outputFolder,
        IProgress<InstallationProgress> progress)
    {
        try
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true,
                StageTitle = (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Unpacking") ??
                              "Unpacking {0}...").Replace("{0}", Path.GetFileName(sourceFile))
            });

            // dark.exe {sourceFile} -x {outDir}
            var procStart = new ProcessStartInfo
            {
                FileName = Path.Combine(wixPath, "dark.exe"),
                WorkingDirectory = (await GetTempDirectory()).Path,
                Arguments = $"\"{sourceFile}\" -x \"{outputFolder}\"",
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,

                // Verbose error handling
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = false,
                UseShellExecute = false,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            var proc = Process.Start(procStart);
            // Redirecting process output so that we can log what happened
            var stdout = new StringBuilder();
            var stderr = new StringBuilder();

            proc!.OutputDataReceived += (_, args) =>
            {
                if (args.Data != null)
                    stdout.AppendLine(args.Data);
            };
            proc.ErrorDataReceived += (_, args) =>
            {
                if (args.Data != null)
                    stderr.AppendLine(args.Data);
            };

            proc.BeginErrorReadLine();
            proc.BeginOutputReadLine();
            var hasExited = proc.WaitForExit(60000);

            // https://github.com/wixtoolset/wix3/blob/6b461364c40e6d1c487043cd0eae7c1a3d15968c/src/tools/dark/dark.cs#L54
            // Exit codes for DARK:
            // 
            // 0 - Success
            // 1 - Error
            // Just in case
            if (!hasExited)
            {
                // WTF
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true, StageTitle =
                        Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Dark/Error/Timeout") ??
                        "Failed to execute dark.exe in the allocated time!"
                });

                proc.Kill();
            }

            if (proc.ExitCode == 1)
            {
                // Assume WiX failed
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true, StageTitle =
                        (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Dark/Error/Result") ??
                         "Dark.exe exited with error code: {0}").Replace("{0}", proc.ExitCode.ToString())
                });

                return false;
            }
        }
        catch (Exception e)
        {
            progress.Report(new InstallationProgress
            {
                IsIndeterminate = true, StageTitle =
                    (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Exceptions/Other") ??
                     "Exception: {0}").Replace("{0}", e.Message)
            });

            return false;
        }

        return true;
    }

    private bool InstallFiles(IEnumerable<string> files, IProgress<InstallationProgress> progress)
    {
        // Execute each install
        foreach (var installFile in files)
            try
            {
                // msi /qn /norestart
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true, StageTitle =
                        (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Installing") ??
                         "Installing {0}...").Replace("{0}", Path.GetFileName(installFile))
                });

                var msiExecutableStart = new ProcessStartInfo
                {
                    FileName = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                        @"System32\msiexec.exe"),
                    WorkingDirectory = Directory.GetParent(installFile)!.FullName,
                    Arguments = $"/i {installFile} /quiet /qn /norestart ALLUSERS=1",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                var msiExecutable = Process.Start(msiExecutableStart);
                msiExecutable!.WaitForExit(60000);
            }
            catch (Exception e)
            {
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true, StageTitle =
                        (Host?.RequestLocalizedString("/Plugins/KinectOne/Stages/Exceptions/Other") ??
                         "Exception: {0}").Replace("{0}", e.Message)
                });

                return false;
            }

        return true;
    }
}

public static class RestExtensions
{
    public static Task<byte[]> ExecuteDownloadDataAsync(this RestClient client, string baseUrl, RestRequest request)
    {
        client.Options.BaseUrl = new Uri(baseUrl);
        return client.DownloadDataAsync(request);
    }

    public static Task<Stream> ExecuteDownloadStreamAsync(this RestClient client, string baseUrl, RestRequest request)
    {
        client.Options.BaseUrl = new Uri(baseUrl);
        return client.DownloadStreamAsync(request);
    }
}

public static class StreamExtensions
{
    public static async Task CopyToWithProgressAsync(this Stream source,
        Stream destination, Action<long> progress = null, int bufferSize = 10240)
    {
        var buffer = new byte[bufferSize];
        var total = 0L;
        int amtRead;

        do
        {
            amtRead = 0;
            while (amtRead < bufferSize)
            {
                var numBytes = await source.ReadAsync(
                    buffer, amtRead, bufferSize - amtRead);
                if (numBytes == 0) break;
                amtRead += numBytes;
            }

            total += amtRead;
            await destination.WriteAsync(buffer, 0, amtRead);
            progress?.Invoke(total);
        } while (amtRead == bufferSize);
    }
}