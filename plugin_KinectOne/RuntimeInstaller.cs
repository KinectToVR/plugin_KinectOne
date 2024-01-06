using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Windows.Storage;
using Amethyst.Plugins.Contract;
using Microsoft.UI.Xaml.Controls;
using RestSharp;
using Microsoft.UI.Xaml.Markup;
using Microsoft.UI.Xaml.Media;

namespace plugin_KinectOne;

internal class SetupData : ICoreSetupData
{
    public object PluginIcon => new PathIcon
    {
        Data = (Geometry)XamlBindingHelper.ConvertValue(typeof(Geometry),
            "M45.26,18.3V15.93H69.51V1.1H0V16H24.25v2.37H0v5.25H69.51V18.3ZM9.36,13.19A4.63,4.63,0,0,1,4.65,8.45a4.61,4.61,0,0,1,4.6-4.67,4.71,4.71,0,1,1,.11,9.41Z")
    };

    public string GroupName => "kinect";
    public Type PluginType => typeof(ITrackingDevice);
}

internal class RuntimeInstaller : IDependencyInstaller
{
    public IDependencyInstaller.ILocalizationHost Host { get; set; }

    public List<IDependency> ListDependencies()
    {
        return new List<IDependency>
        {
            new KinectRuntime
            {
                Host = Host,
                Name = Host?.RequestLocalizedString("/Plugins/KinectOne/Dependencies/Runtime/Name") ??
                       "Kinect for Xbox One Runtime"
            }
        };
    }
}

internal class KinectRuntime : IDependency
{
    public IDependencyInstaller.ILocalizationHost Host { get; set; }

    public string Name { get; set; }
    public bool IsMandatory => true;

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

    public async Task<bool> Install(IProgress<InstallationProgress> progress, CancellationToken cancellationToken)
    {
        // Amethyst will handle this exception for us anyway
        cancellationToken.ThrowIfCancellationRequested();
        var paths = new[]
        {
            Path.Join(Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.FullName,
                "Assets", "Resources", "Dependencies", "KinectRuntime-x64.msi")
        };

        // Copy to temp if amethyst is packaged
        // ReSharper disable once InvertIf
        if (PackageUtils.IsAmethystPackaged)
        {
            // Create a shared folder with the dependencies
            var dependenciesFolder = await ApplicationData.Current.TemporaryFolder.CreateFolderAsync(
                Guid.NewGuid().ToString().ToUpper(), CreationCollisionOption.OpenIfExists);

            // Copy all driver files to Amethyst's local data folder
            new DirectoryInfo(Path.Join(Directory.GetParent(Assembly.GetExecutingAssembly().Location)!.FullName,
                    "Assets", "Resources", "Dependencies"))
                .CopyToFolder(dependenciesFolder.Path);

            // Update the installation paths
            paths = new[]
            {
                Path.Join(dependenciesFolder.Path, "KinectRuntime-x64.msi")
            };
        }

        // Finally install the packages
        return InstallFiles(paths, progress, cancellationToken);
    }

    private bool InstallFiles(IEnumerable<string> files,
        IProgress<InstallationProgress> progress, CancellationToken cancellationToken)
    {
        // Amethyst will handle this exception for us anyway
        cancellationToken.ThrowIfCancellationRequested();

        // Execute each install
        foreach (var installFile in files)
            try
            {
                // msi /qn /norestart
                progress.Report(new InstallationProgress
                {
                    IsIndeterminate = true,
                    StageTitle =
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
                    IsIndeterminate = true,
                    StageTitle =
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
        Stream destination, CancellationToken cancellationToken,
        Action<long> progress = null, int bufferSize = 10240)
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
                    buffer, amtRead, bufferSize - amtRead, cancellationToken);
                if (numBytes == 0) break;
                amtRead += numBytes;
            }

            total += amtRead;
            await destination.WriteAsync(buffer, 0, amtRead, cancellationToken);
            progress?.Invoke(total);
        } while (amtRead == bufferSize);
    }
}