#pragma once

#pragma unmanaged
#include "KinectWrapper.h"
#pragma managed

using namespace System;
using namespace Numerics;
using namespace Collections::Generic;
using namespace ComponentModel::Composition;
using namespace Runtime::InteropServices;

namespace KinectHandler
{
    private enum TrackedJointTypeInternal
    {
        JointHead,
        JointNeck,
        JointSpineShoulder,
        JointShoulderLeft,
        JointElbowLeft,
        JointWristLeft,
        JointHandLeft,
        JointHandTipLeft,
        JointThumbLeft,
        JointShoulderRight,
        JointElbowRight,
        JointWristRight,
        JointHandRight,
        JointHandTipRight,
        JointThumbRight,
        JointSpineMiddle,
        JointSpineWaist,
        JointHipLeft,
        JointKneeLeft,
        JointFootLeft,
        JointFootTipLeft,
        JointHipRight,
        JointKneeRight,
        JointFootRight,
        JointFootTipRight,
        JointManual,
        JointCount
    };

    public ref class KinectJoint sealed
    {
    public:
        KinectJoint(const int role)
        {
            JointRole = role;
        }

        property Vector3 Position;
        property Quaternion Orientation;

        property int TrackingState;
        property int JointRole;
    };

    delegate void FunctionToCallDelegate();

    public ref class KinectHandler
    {
    private:
        KinectWrapper* kinect_;
        FunctionToCallDelegate^ function_;
        array<BYTE>^ managed_buffer_ = nullptr;

    public:
        KinectHandler() : kinect_(new KinectWrapper())
        {
            function_ = gcnew FunctionToCallDelegate(this, &KinectHandler::StatusChangedHandler);
            [[maybe_unused]] pin_ptr<FunctionToCallDelegate^> tmp = &function_; // Pin the function delegate

            status_changed_event = static_cast<void(__cdecl*)()>(
                Marshal::GetFunctionPointerForDelegate(function_).ToPointer());
        }

        virtual void StatusChangedHandler()
        {
            // implemented in the c# handler
        }

        array<BYTE>^ GetImageBuffer()
        {
            if (!IsInitialized || !kinect_->camera_enabled()) return __nullptr;
            const auto& [unmanagedBuffer, size] = kinect_->color_buffer();
            if (!unmanagedBuffer || size <= 0) return __nullptr;

            if (managed_buffer_ == nullptr || managed_buffer_->Length != size)
                managed_buffer_ = gcnew array<BYTE>(size); // allocate once

            Marshal::Copy(IntPtr(unmanagedBuffer), managed_buffer_, 0, size);
            return managed_buffer_;
        }

        List<KinectJoint^>^ GetTrackedKinectJoints()
        {
            if (!IsInitialized) return gcnew List<KinectJoint^>;

            const auto& positions = kinect_->skeleton_positions();
            const auto& orientations = kinect_->bone_orientations();

            auto trackedKinectJoints = gcnew List<KinectJoint^>;
            for (auto v = 0; v < JointCount; v++)
            {
                if (v == JointManual)
                    continue; // Skip unsupported joints

                auto joint = gcnew KinectJoint(v);

                joint->TrackingState =
                    positions[kinect_->KinectJointType(v)].TrackingState;

                joint->Position = Vector3(
                    positions[kinect_->KinectJointType(v)].Position.X,
                    positions[kinect_->KinectJointType(v)].Position.Y,
                    positions[kinect_->KinectJointType(v)].Position.Z);

                joint->Orientation = Quaternion(
                    orientations[kinect_->KinectJointType(v)].Orientation.x,
                    orientations[kinect_->KinectJointType(v)].Orientation.y,
                    orientations[kinect_->KinectJointType(v)].Orientation.z,
                    orientations[kinect_->KinectJointType(v)].Orientation.w);

                trackedKinectJoints->Add(joint);
            }

            return trackedKinectJoints;
        }

        property bool LeftHandClosed
        {
            bool get() { return kinect_->left_hand_state(); }
        }

        property bool RightHandClosed
        {
            bool get() { return kinect_->right_hand_state(); }
        }

        property bool IsInitialized
        {
            bool get() { return kinect_->is_initialized(); }
        }

        property bool IsSkeletonTracked
        {
            bool get() { return kinect_->skeleton_tracked(); }
        }

        property int DeviceStatus
        {
            int get() { return kinect_->status_result(); }
        }

        property bool IsCameraEnabled
        {
            bool get() { return kinect_->camera_enabled(); }
            void set(const bool value) { kinect_->camera_enabled(value); }
        }

        property int CameraImageWidth
        {
            int get() { return kinect_->CameraImageSize().first; }
        }

        property int CameraImageHeight
        {
            int get() { return kinect_->CameraImageSize().second; }
        }

        Drawing::Size MapCoordinate(Vector3 position)
        {
            if (!IsInitialized) return Drawing::Size::Empty;
            const auto& [width, height] =
                kinect_->MapCoordinate(CameraSpacePoint{position.X, position.Y, position.Z});

            return Drawing::Size(width, height);
        }

        int InitializeKinect()
        {
            return kinect_->initialize();
        }

        int ShutdownKinect()
        {
            return kinect_->shutdown();
        }
    };
}
