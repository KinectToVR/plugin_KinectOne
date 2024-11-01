// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Drawing;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices.WindowsRuntime;
using Amethyst.Plugins.Contract;
using Microsoft.UI.Xaml.Media.Imaging;

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
    public WriteableBitmap CameraImage { get; set; }
    public static IAmethystHost HostStatic { get; set; }

    private readonly GestureDetector
        _pauseDetectorLeft = new(),
        _pauseDetectorRight = new(),
        _pointDetectorLeft = new(),
        _pointDetectorRight = new();

    public ObservableCollection<TrackedJoint> TrackedJoints { get; } =
        // Prepend all supported joints to the joints list
        new(Enum.GetValues<TrackedJointType>()
            .Where(x => x is not TrackedJointType.JointManual)
            .Select(x => new TrackedJoint
            {
                Name = HostStatic?.RequestLocalizedString($"/JointsEnum/{x.ToString()}") ?? x.ToString(),
                Role = x,
                SupportedInputActions = x switch
                {
                    TrackedJointType.JointHandLeft =>
                    [
                        new KeyInputAction<bool>
                        {
                            Name = "Left Pause", Description = "Left hand pause gesture",
                            Guid = "5E4680F9-F232-4EA1-AE12-E96F7F8E0CC1", GetHost = () => HostStatic
                        },
                        new KeyInputAction<bool>
                        {
                            Name = "Left Point", Description = "Left hand point gesture",
                            Guid = "8D83B89D-5FBD-4D52-B626-4E90BDD26B08", GetHost = () => HostStatic
                        },
                        new KeyInputAction<bool>
                        {
                            Name = "Left Grab", Description = "Left hand grab gesture",
                            Guid = "E383258F-5918-4F1C-BC66-7325DB1F07E8", GetHost = () => HostStatic
                        }
                    ],
                    TrackedJointType.JointHandRight =>
                    [
                        new KeyInputAction<bool>
                        {
                            Name = "Right Pause", Description = "Right hand pause gesture",
                            Guid = "B8389FA6-75EF-4509-AEC2-1758AFE41D95", GetHost = () => HostStatic
                        },
                        new KeyInputAction<bool>
                        {
                            Name = "Right Point", Description = "Right hand point gesture",
                            Guid = "C58EBCFE-0DF5-40FD-ABC1-06B415FA51BE", GetHost = () => HostStatic
                        },
                        new KeyInputAction<bool>
                        {
                            Name = "Right Grab", Description = "Right hand grab gesture",
                            Guid = "801336BE-5BD5-4881-A390-D57D958592EF", GetHost = () => HostStatic
                        }
                    ],
                    _ => []
                }
            }));

    public string DeviceStatusString => PluginLoaded
        ? DeviceStatus switch
        {
            0 => Host.RequestLocalizedString("/Plugins/KinectOne/Statuses/Success"),
            1 => Host.RequestLocalizedString("/Plugins/KinectOne/Statuses/NotAvailable"),
            _ => $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, although we can't tell what."
        }
        : $"Undefined: {DeviceStatus}\nE_UNDEFINED\nSomething weird has happened, although we can't tell what.";

    public Uri ErrorDocsUri => new($"https://docs.k2vr.tech/{Host?.DocsLanguageCode ?? "en"}/one/troubleshooting/");

    public void OnLoad()
    {
        // Backup the plugin host
        HostStatic = Host;
        PluginLoaded = true;

        CameraImage = new WriteableBitmap(CameraImageWidth, CameraImageHeight);

        try
        {
            // Re-generate joint names
            lock (Host.UpdateThreadLock)
            {
                for (var i = 0; i < TrackedJoints.Count; i++)
                    TrackedJoints[i] = TrackedJoints[i].WithName(Host?.RequestLocalizedString(
                        $"/JointsEnum/{TrackedJoints[i].Role.ToString()}") ?? TrackedJoints[i].Role.ToString());
            }
        }
        catch (Exception e)
        {
            Host?.Log($"Error setting joint names! Message: {e.Message}", LogSeverity.Error);
        }
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

        // Update camera feed
        if (IsCameraEnabled)
            CameraImage.DispatcherQueue.TryEnqueue(async () =>
            {
                var buffer = GetImageBuffer(); // Read from Kinect
                if (buffer is null || buffer.Length <= 0) return;
                await CameraImage.PixelBuffer.AsStream().WriteAsync(buffer);
                CameraImage.Invalidate(); // Enqueue for preview refresh
            });


        // Update gestures
        if (trackedJoints.Count != 25) return;

        try
        {
            var shoulderLeft = trackedJoints[(int)TrackedJointType.JointShoulderLeft].Position;
            var shoulderRight = trackedJoints[(int)TrackedJointType.JointShoulderRight].Position;
            var elbowLeft = trackedJoints[(int)TrackedJointType.JointElbowLeft].Position;
            var elbowRight = trackedJoints[(int)TrackedJointType.JointElbowRight].Position;
            var handLeft = trackedJoints[(int)TrackedJointType.JointWristLeft].Position;
            var handRight = trackedJoints[(int)TrackedJointType.JointWristRight].Position;

            // >0.9f when elbow is not bent and the arm is straight : LEFT
            var armDotLeft = Vector3.Dot(
                Vector3.Normalize(elbowLeft - shoulderLeft),
                Vector3.Normalize(handLeft - elbowLeft));

            // >0.9f when the arm is pointing down : LEFT
            var armDownDotLeft = Vector3.Dot(
                new Vector3(0.0f, -1.0f, 0.0f),
                Vector3.Normalize(handLeft - elbowLeft));

            // >0.4f <0.6f when the arm is slightly tilted sideways : RIGHT
            var armTiltDotLeft = Vector3.Dot(
                Vector3.Normalize(shoulderLeft - shoulderRight),
                Vector3.Normalize(handLeft - elbowLeft));

            // >0.9f when elbow is not bent and the arm is straight : LEFT
            var armDotRight = Vector3.Dot(
                Vector3.Normalize(elbowRight - shoulderRight),
                Vector3.Normalize(handRight - elbowRight));

            // >0.9f when the arm is pointing down : RIGHT
            var armDownDotRight = Vector3.Dot(
                new Vector3(0.0f, -1.0f, 0.0f),
                Vector3.Normalize(handRight - elbowRight));

            // >0.4f <0.6f when the arm is slightly tilted sideways : RIGHT
            var armTiltDotRight = Vector3.Dot(
                Vector3.Normalize(shoulderRight - shoulderLeft),
                Vector3.Normalize(handRight - elbowRight));

            /* Trigger the detected gestures */

            if (TrackedJoints[(int)TrackedJointType.JointHandLeft].SupportedInputActions.IsUsed(0, out var pauseActionLeft))
                Host.ReceiveKeyInput(pauseActionLeft, _pauseDetectorLeft.Update(armDotLeft > 0.9f && armTiltDotLeft is > 0.4f and < 0.7f));

            if (TrackedJoints[(int)TrackedJointType.JointHandRight].SupportedInputActions.IsUsed(0, out var pauseActionRight))
                Host.ReceiveKeyInput(pauseActionRight, _pauseDetectorRight.Update(armDotRight > 0.9f && armTiltDotRight is > 0.4f and < 0.7f));

            if (TrackedJoints[(int)TrackedJointType.JointHandLeft].SupportedInputActions.IsUsed(1, out var pointActionLeft))
                Host.ReceiveKeyInput(pointActionLeft, _pointDetectorLeft
                    .Update(armDotLeft > 0.9f && armTiltDotLeft is > -0.5f and < 0.5f && armDownDotLeft is > -0.3f and < 0.7f));

            if (TrackedJoints[(int)TrackedJointType.JointHandRight].SupportedInputActions.IsUsed(1, out var pointActionRight))
                Host.ReceiveKeyInput(pointActionRight, _pointDetectorRight
                    .Update(armDotRight > 0.9f && armTiltDotRight is > -0.5f and < 0.5f && armDownDotRight is > -0.3f and < 0.7f));

            if (TrackedJoints[(int)TrackedJointType.JointHandLeft].SupportedInputActions.IsUsed(2, out var grabActionLeft))
                Host.ReceiveKeyInput(grabActionLeft, LeftHandClosed);

            if (TrackedJoints[(int)TrackedJointType.JointHandRight].SupportedInputActions.IsUsed(2, out var grabActionRight))
                Host.ReceiveKeyInput(grabActionRight, RightHandClosed);

            /* Trigger the detected gestures */
        }
        catch (Exception ex)
        {
            Host?.Log(ex);
        }
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

    public Func<BitmapSource> GetCameraImage => () => CameraImage;
    public Func<bool> GetIsCameraEnabled => () => IsCameraEnabled;
    public Action<bool> SetIsCameraEnabled => value => IsCameraEnabled = value;
    public Func<Vector3, Size> MapCoordinateDelegate => MapCoordinate;
}

internal static class Utils
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

    public static T At<T>(this SortedSet<T> set, int at)
    {
        return set.ElementAt(at);
    }

    public static bool At<T>(this SortedSet<T> set, int at, out T result)
    {
        try
        {
            result = set.ElementAt(at);
        }
        catch
        {
            result = default;
            return false;
        }

        return true;
    }

    public static bool IsUsed(this SortedSet<IKeyInputAction> set, int at)
    {
        return set.At(at, out var action) && (KinectOne.HostStatic?.CheckInputActionIsUsed(action) ?? false);
    }

    public static bool IsUsed(this SortedSet<IKeyInputAction> set, int at, out IKeyInputAction action)
    {
        return set.At(at, out action) && (KinectOne.HostStatic?.CheckInputActionIsUsed(action) ?? false);
    }

    public static TrackedJoint WithName(this TrackedJoint joint, string name)
    {
        return new TrackedJoint
        {
            Name = name,
            Role = joint.Role,
            Acceleration = joint.Acceleration,
            AngularAcceleration = joint.AngularAcceleration,
            AngularVelocity = joint.AngularVelocity,
            Orientation = joint.Orientation,
            Position = joint.Position,
            SupportedInputActions = joint.SupportedInputActions,
            TrackingState = joint.TrackingState,
            Velocity = joint.Velocity
        };
    }
}