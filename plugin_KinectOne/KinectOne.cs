// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.Composition;
using System.Linq;
using System.Numerics;
using System.Threading;
using Amethyst.Plugins.Contract;
using Microsoft.Kinect;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace plugin_KinectOne;

[Export(typeof(ITrackingDevice))]
[ExportMetadata("Name", "Xbox One Kinect")]
[ExportMetadata("Guid", "K2VRTEAM-AME2-APII-DVCE-DVCEKINECTV2")]
[ExportMetadata("Publisher", "K2VR Team")]
[ExportMetadata("Version", "1.0.0.0")]
[ExportMetadata("Website", "https://github.com/KinectToVR/plugin_KinectOne")]
[ExportMetadata("DependencyLink", "https://docs.k2vr.tech/{0}/one/setup/")]
[ExportMetadata("DependencySource",
    "https://download.microsoft.com/download/A/7/4/A74239EB-22C2-45A1-996C-2F8E564B28ED/KinectRuntime-v2.0_1409-Setup.exe")]
[ExportMetadata("DependencyInstaller", typeof(RuntimeInstaller))]
public class KinectOne : ITrackingDevice
{
    private static readonly SortedDictionary<TrackedJointType, JointType> KinectJointTypeDictionary = new()
    {
        { TrackedJointType.JointHead, JointType.Head },
        { TrackedJointType.JointNeck, JointType.Neck },
        { TrackedJointType.JointSpineShoulder, JointType.SpineShoulder },
        { TrackedJointType.JointShoulderLeft, JointType.ShoulderLeft },
        { TrackedJointType.JointElbowLeft, JointType.ElbowLeft },
        { TrackedJointType.JointWristLeft, JointType.WristLeft },
        { TrackedJointType.JointHandLeft, JointType.HandLeft },
        { TrackedJointType.JointHandTipLeft, JointType.HandTipLeft },
        { TrackedJointType.JointThumbLeft, JointType.ThumbLeft },
        { TrackedJointType.JointShoulderRight, JointType.ShoulderRight },
        { TrackedJointType.JointElbowRight, JointType.ElbowRight },
        { TrackedJointType.JointWristRight, JointType.WristRight },
        { TrackedJointType.JointHandRight, JointType.HandRight },
        { TrackedJointType.JointHandTipRight, JointType.HandTipRight },
        { TrackedJointType.JointThumbRight, JointType.ThumbRight },
        { TrackedJointType.JointSpineMiddle, JointType.SpineMid },
        { TrackedJointType.JointSpineWaist, JointType.SpineBase },
        { TrackedJointType.JointHipLeft, JointType.HipLeft },
        { TrackedJointType.JointKneeLeft, JointType.KneeLeft },
        { TrackedJointType.JointFootLeft, JointType.AnkleLeft },
        { TrackedJointType.JointFootTipLeft, JointType.FootLeft },
        { TrackedJointType.JointHipRight, JointType.HipRight },
        { TrackedJointType.JointKneeRight, JointType.KneeRight },
        { TrackedJointType.JointFootRight, JointType.AnkleRight },
        { TrackedJointType.JointFootTipRight, JointType.FootRight }
    };

    [Import(typeof(IAmethystHost))] private IAmethystHost Host { get; set; }

    private KinectSensor KinectSensor { get; set; }
    private BodyFrameReader BodyFrameReader { get; set; }
    private Body[] Bodies { get; set; }
    private bool PluginLoaded { get; set; }

    public bool IsPositionFilterBlockingEnabled => false;
    public bool IsPhysicsOverrideEnabled => false;
    public bool IsSelfUpdateEnabled => true;
    public bool IsFlipSupported => true;
    public bool IsAppOrientationSupported => true;
    public bool IsSettingsDaemonSupported => false;
    public object SettingsInterfaceRoot => null;

    public bool IsInitialized { get; private set; }

    public bool IsSkeletonTracked { get; private set; }

    public int DeviceStatus => KinectSensor?.IsAvailable ?? false ? 0 : 1;

    public ObservableCollection<TrackedJoint> TrackedJoints { get; } =
        // Prepend all supported joints to the joints list
        new(Enum.GetValues<TrackedJointType>().Where(x => x != TrackedJointType.JointManual)
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
        IsInitialized = InitKinect();
        Host.Log($"Tried to initialize the Kinect sensor with status: {DeviceStatusString}");

        // Try to start the stream
        InitializeSkeleton();
    }

    public void Shutdown()
    {
        // BodyFrameReader is IDisposable
        BodyFrameReader?.Dispose();
        BodyFrameReader = null;

        // Close the kinect sensor
        KinectSensor?.Close();
        KinectSensor = null;

        // Mark as not initialized
        IsInitialized = false;
    }

    public void Update()
    {
        // ignored
    }

    public void SignalJoint(int jointId)
    {
        // ignored
    }

    private bool InitKinect()
    {
        if ((KinectSensor = KinectSensor.GetDefault()) is null) return false;

        try
        {
            // Try to open the kinect sensor
            KinectSensor.Open();

            // Necessary to allow kinect to become available behind the scenes
            Thread.Sleep(2000);

            // Register a watchdog (remove, add)
            KinectSensor.IsAvailableChanged -= StatusChangedHandler;
            KinectSensor.IsAvailableChanged += StatusChangedHandler;

            // Check the status and return
            return KinectSensor.IsAvailable;
        }
        catch (Exception e)
        {
            Host.Log($"Failed to open the Kinect sensor! Message: {e.Message}");
            return false;
        }
    }

    private void StatusChangedHandler(object o, IsAvailableChangedEventArgs isAvailableChangedEventArgs)
    {
        // Make AME refresh our plugin
        Host?.RefreshStatusInterface();
    }

    private void InitializeSkeleton()
    {
        BodyFrameReader?.Dispose();
        BodyFrameReader = KinectSensor.BodyFrameSource.OpenReader();

        if (BodyFrameReader is null) return;

        // Register a watchdog (remove, add)
        BodyFrameReader.FrameArrived -= OnBodyFrameArrivedHandler;
        BodyFrameReader.FrameArrived += OnBodyFrameArrivedHandler;
    }

    private void OnBodyFrameArrivedHandler(object _, BodyFrameArrivedEventArgs args)
    {
        var dataReceived = false;
        using var bodyFrame = args.FrameReference.AcquireFrame();
        if (bodyFrame != null)
        {
            Bodies ??= new Body[bodyFrame.BodyCount];

            // The first time GetAndRefreshBodyData is called, Kinect will allocate each Body in the array.
            // As long as those body objects are not disposed and not set to null in the array,
            // those body objects will be re-used.
            bodyFrame.GetAndRefreshBodyData(Bodies);
            dataReceived = true;
        }

        // Validate the result from the sensor
        if (!dataReceived || Bodies.Length <= 0)
        {
            IsSkeletonTracked = false;
            return; // Give up this time
        }

        // Check if any body is tracked, copy the status
        IsSkeletonTracked = Bodies.Any(x => x.IsTracked);
        if (!IsSkeletonTracked) return; // Give up

        // Get the first tracked body, decompose
        var body = Bodies.First(x => x.IsTracked);
        var jointPositions = body.Joints;
        var jointOrientations = body.JointOrientations;

        // Copy positions, orientations and states from the sensor
        // We should be able to address our joints by [] because
        // they're prepended via Enum.GetValues<TrackedJointType>
        foreach (var (appJoint, kinectJoint) in KinectJointTypeDictionary)
        {
            var tracker = TrackedJoints[(int)appJoint];
            (tracker.Position, tracker.Orientation, tracker.TrackingState) = (
                jointPositions.First(x => x.Key == kinectJoint).Value.PoseVector(),
                jointOrientations.First(x => x.Key == kinectJoint).Value.OrientationQuaternion(),
                (TrackedJointState)jointPositions.First(x => x.Key == kinectJoint).Value.TrackingState);
        }

        //// Fix orientations: knees and elbows appear sideways (left)
        //TrackedJoints[(int)TrackedJointType.JointAnkleLeft].JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(-Math.PI / 3.0), 0f);
        //TrackedJoints[(int)TrackedJointType.JointKneeLeft].JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(-Math.PI / 3.0), 0f);

        //// Fix orientations: knees and elbows appear sideways (right)
        //TrackedJoints[(int)TrackedJointType.JointAnkleRight].JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(Math.PI / 3.0), 0f);
        //TrackedJoints[(int)TrackedJointType.JointKneeRight].JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(Math.PI / 3.0), 0f);

        //foreach (var (appJoint, kinectJoint) in KinectJointTypeDictionary)
        //{
        //    var tracker = TrackedJoints.First(x => x.Role == appJoint);
        //    (tracker.JointPosition, tracker.JointOrientation, tracker.TrackingState) = (
        //        jointPositions.First(x => x.Key == kinectJoint).Value.PoseVector(),
        //        jointOrientations.First(x => x.Key == kinectJoint).Value.OrientationQuaternion(),
        //        KinectJointStateDictionary[
        //            jointPositions.First(x => x.Key == kinectJoint).Value.TrackingState]);
        //}

        //TrackedJoints.First(x => x.Role is TrackedJointType.JointAnkleLeft).JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(-Math.PI / 3.0), 0f);
        //TrackedJoints.First(x => x.Role is TrackedJointType.JointKneeLeft).JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(-Math.PI / 3.0), 0f);

        //TrackedJoints.First(x => x.Role is TrackedJointType.JointAnkleRight).JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(Math.PI / 3.0), 0f);
        //TrackedJoints.First(x => x.Role is TrackedJointType.JointKneeRight).JointOrientation *=
        //    Quaternion.CreateFromYawPitchRoll(0f, (float)(Math.PI / 3.0), 0f);
    }
}

public static class JointExtensions
{
    public static Vector3 PoseVector(this Joint joint)
    {
        return new Vector3(joint.Position.X, joint.Position.Y, joint.Position.Z);
    }

    public static Quaternion OrientationQuaternion(this JointOrientation joint)
    {
        return new Quaternion(joint.Orientation.X, joint.Orientation.Y, joint.Orientation.Z, joint.Orientation.W);
    }
}