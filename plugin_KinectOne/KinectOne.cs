// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See LICENSE in the project root for license information.

using System;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Linq;
using System.Numerics;
using Amethyst.Plugins.Contract;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace plugin_KinectOne;

[Export(typeof(ITrackingDevice))]
[ExportMetadata("Name", "Xbox One Kinect")]
[ExportMetadata("Guid", "K2VRTEAM-AME2-APII-DVCE-DVCEKINECTV2")]
[ExportMetadata("Publisher", "K2VR Team")]
[ExportMetadata("Version", "1.0.0.1")]
[ExportMetadata("Website", "https://github.com/KimihikoAkayasaki/plugin_KinectOne")]
[ExportMetadata("DependencyLink", "https://docs.k2vr.tech/{0}/one/setup/")]
[ExportMetadata("DependencySource",
    "https://download.microsoft.com/download/A/7/4/A74239EB-22C2-45A1-996C-2F8E564B28ED/KinectRuntime-v2.0_1409-Setup.exe")]
[ExportMetadata("DependencyInstaller", typeof(RuntimeInstaller))]
[ExportMetadata("CoreSetupData", typeof(SetupData))]
public class KinectOne : KinectHandler.KinectHandler, ITrackingDevice
{
    [Import(typeof(IAmethystHost))] private IAmethystHost Host { get; set; }

    private bool PluginLoaded { get; set; }
    public bool IsPositionFilterBlockingEnabled => false;
    public bool IsPhysicsOverrideEnabled => false;
    public bool IsSelfUpdateEnabled => false;
    public bool IsFlipSupported => true;
    public bool IsAppOrientationSupported => true;
    public object SettingsInterfaceRoot => null;

    public ObservableCollection<TrackedJoint> TrackedJoints { get; } =
        // Prepend all supported joints to the joints list
        new(Enum.GetValues<TrackedJointType>()
            .Where(x => x is not TrackedJointType.JointManual)
            .Select(x => new TrackedJoint { Name = x.ToString(), Role = x }));

    public string DeviceStatusString => PluginLoaded
        ? DeviceStatus switch
        {
            0 => Host.RequestLocalizedString("/Plugins/KinectOne/Statuses/Success"),
            1 => Host.RequestLocalizedString("/Plugins/KinectOne/Statuses/NotAvailable"),
            _ => $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, though we can't tell what."
        }
        : $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, though we can't tell what.";

    public Uri ErrorDocsUri => new($"https://docs.k2vr.tech/{Host?.DocsLanguageCode ?? "en"}/one/troubleshooting/");

    public void OnLoad()
    {
        PluginLoaded = true;
    }

    public void Initialize()
    {
        switch (InitializeKinect())
        {
            case 0:
                Host.Log($"Tried to initialize the Kinect sensor with status: {DeviceStatusString}");
                break;
            case 1:
                Host.Log($"Couldn't initialize the Kinect sensor! Status: {DeviceStatusString}", LogSeverity.Warning);
                break;
            default:
                Host.Log("Tried to initialize the Kinect, but a native exception occurred!", LogSeverity.Error);
                break;
        }
    }

    public void Shutdown()
    {
        switch (ShutdownKinect())
        {
            case 0:
                Host.Log($"Tried to shutdown the Kinect sensor with status: {DeviceStatusString}");
                break;
            case 1:
                Host.Log($"Kinect sensor is already shut down! Status: {DeviceStatusString}", LogSeverity.Warning);
                break;
            case -2:
                Host.Log("Tried to shutdown the Kinect sensor, but a SEH exception occurred!", LogSeverity.Error);
                break;
            default:
                Host.Log("Tried to shutdown the Kinect sensor, but a native exception occurred!", LogSeverity.Error);
                break;
        }
    }

    public void Update()
    {
        var trackedJoints = GetTrackedKinectJoints();
        trackedJoints.ForEach(x =>
        {
            TrackedJoints[trackedJoints.IndexOf(x)].TrackingState =
                (TrackedJointState)x.TrackingState;

            TrackedJoints[trackedJoints.IndexOf(x)].Position = x.Position.Safe();
            TrackedJoints[trackedJoints.IndexOf(x)].Orientation = x.Orientation.Safe();
        });
    }

    public void SignalJoint(int jointId)
    {
        // ignored
    }

    public override void StatusChangedHandler()
    {
        // The Kinect sensor requested a refresh
        InitializeKinect();

        // Request a refresh of the status UI
        Host?.RefreshStatusInterface();
    }
}

internal static class PoseUtils
{
    public static Quaternion Safe(this Quaternion q)
    {
        return (q.X is 0 && q.Y is 0 && q.Z is 0 && q.W is 0) ||
               float.IsNaN(q.X) || float.IsNaN(q.Y) || float.IsNaN(q.Z) || float.IsNaN(q.W)
            ? Quaternion.Identity // Return a placeholder quaternion
            : q; // If everything is fine, return the actual orientation
    }

    public static Vector3 Safe(this Vector3 v)
    {
        return float.IsNaN(v.X) || float.IsNaN(v.Y) || float.IsNaN(v.Z)
            ? Vector3.Zero // Return a placeholder position vector
            : v; // If everything is fine, return the actual orientation
    }
}