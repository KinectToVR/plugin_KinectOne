#pragma once
#include <Windows.h>
#include <Ole2.h>
#include <Kinect.h>
#include <atlbase.h>

#include <algorithm>
#include <iterator>
#include <memory>
#include <thread>
#include <array>
#include <map>
#include <functional>

inline void (*status_changed_event)();

class KinectWrapper
{
    IKinectSensor* kinectSensor = nullptr;
    IBodyFrameReader* bodyFrameReader = nullptr;
    IColorFrameReader* colorFrameReader = nullptr;
    IDepthFrameReader* depthFrameReader = nullptr;
    IMultiSourceFrame* multiFrame = nullptr;
    ICoordinateMapper* coordMapper = nullptr;
    BOOLEAN isTracking = false;

    Joint joints[JointType_Count];
    JointOrientation boneOrientations[JointType_Count];
    IBody* kinectBodies[BODY_COUNT] = {nullptr};

    WAITABLE_HANDLE h_statusChangedEvent;
    WAITABLE_HANDLE h_bodyFrameEvent;
    bool newBodyFrameArrived = false;

    std::array<JointOrientation, JointType_Count> bone_orientations_;
    std::array<Joint, JointType_Count> skeleton_positions_;

    std::unique_ptr<std::thread> updater_thread_;

    inline static bool initialized_ = false;
    bool skeleton_tracked_ = false;

    void updater()
    {
        // Auto-handles failures & etc
        while (true)update();
    }

    void updateSkeletalData()
    {
        if (!bodyFrameReader)return; // Give up already

        IBodyFrame* bodyFrame = nullptr;
        bodyFrameReader->AcquireLatestFrame(&bodyFrame);

        if (!bodyFrame) return;

        bodyFrame->GetAndRefreshBodyData(BODY_COUNT, kinectBodies);
        newBodyFrameArrived = true;
        if (bodyFrame) bodyFrame->Release();

        // We have the frame, now parse it
        for (const auto& i : kinectBodies)
        {
            BOOLEAN isSkeletonTracked = false;
            if (i)i->get_IsTracked(&isSkeletonTracked);

            if (isSkeletonTracked)
            {
                skeleton_tracked_ = isSkeletonTracked;
                i->GetJoints(JointType_Count, joints);
                i->GetJointOrientations(JointType_Count, boneOrientations);

                // Copy joint positions & orientations
                std::copy(std::begin(joints), std::end(joints),
                          skeleton_positions_.begin());
                std::copy(std::begin(boneOrientations), std::end(boneOrientations),
                          bone_orientations_.begin());

                break;
            }
            skeleton_tracked_ = false;
        }
    }

    bool initKinect()
    {
        // Get a working Kinect Sensor
        if (FAILED(GetDefaultKinectSensor(&kinectSensor))) return false;
        if (kinectSensor)
        {
            kinectSensor->get_CoordinateMapper(&coordMapper);
            HRESULT hr_open = kinectSensor->Open();

            // Necessary to allow kinect to become available behind the scenes
            std::this_thread::sleep_for(std::chrono::seconds(2));

            BOOLEAN available = false;
            kinectSensor->get_IsAvailable(&available);

            // Check the sensor (just in case)
            if (FAILED(hr_open) || !available || kinectSensor == nullptr)
            {
                return false;
            }

            // Register a StatusChanged event
            h_statusChangedEvent = (WAITABLE_HANDLE)CreateEvent(nullptr, FALSE, FALSE, nullptr);
            kinectSensor->SubscribeIsAvailableChanged(&h_statusChangedEvent);

            return true;
        }

        return false;
    }

    void initializeSkeleton()
    {
        if (bodyFrameReader)
            bodyFrameReader->Release();

        IBodyFrameSource* bodyFrameSource;
        kinectSensor->get_BodyFrameSource(&bodyFrameSource);
        bodyFrameSource->OpenReader(&bodyFrameReader);

        // Newfangled event based frame capture
        // https://github.com/StevenHickson/PCL_Kinect2SDK/blob/master/src/Microsoft_grabber2.cpp
        h_bodyFrameEvent = (WAITABLE_HANDLE)CreateEvent(nullptr, FALSE, FALSE, nullptr);
        bodyFrameReader->SubscribeFrameArrived(&h_bodyFrameEvent);
        if (bodyFrameSource) bodyFrameSource->Release();
    }

    void terminateSkeleton()
    {
        if (!bodyFrameReader)return; // No need to do anything
        if (FAILED(bodyFrameReader->UnsubscribeFrameArrived(h_bodyFrameEvent)))
        {
            throw std::exception("Couldn't unsubscribe frame!");
        }
        __try
        {
            CloseHandle((HANDLE)h_bodyFrameEvent);
            bodyFrameReader->Release();
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            // ignored
        }
        h_bodyFrameEvent = NULL;
        bodyFrameReader = nullptr;
    }

    HRESULT kinect_status_result()
    {
        BOOLEAN avail;
        if (kinectSensor)
        {
            kinectSensor->get_IsAvailable(&avail);
            if (avail)
                return S_OK;
        }

        return S_FALSE;
    }

    enum JointType
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
        JointManual
    };

    std::map<JointType, _JointType> KinectJointTypeDictionary
    {
        {JointHead, JointType_Head},
        {JointNeck, JointType_Neck},
        {JointSpineShoulder, JointType_SpineShoulder},
        {JointShoulderLeft, JointType_ShoulderLeft},
        {JointElbowLeft, JointType_ElbowLeft},
        {JointWristLeft, JointType_WristLeft},
        {JointHandLeft, JointType_HandLeft},
        {JointHandTipLeft, JointType_HandTipLeft},
        {JointThumbLeft, JointType_ThumbLeft},
        {JointShoulderRight, JointType_ShoulderRight},
        {JointElbowRight, JointType_ElbowRight},
        {JointWristRight, JointType_WristRight},
        {JointHandRight, JointType_HandRight},
        {JointHandTipRight, JointType_HandTipRight},
        {JointThumbRight, JointType_ThumbRight},
        {JointSpineMiddle, JointType_SpineMid},
        {JointSpineWaist, JointType_SpineBase},
        {JointHipLeft, JointType_HipLeft},
        {JointKneeLeft, JointType_KneeLeft},
        {JointFootLeft, JointType_FootLeft},
        {JointFootTipLeft, JointType_FootLeft},
        {JointHipRight, JointType_HipRight},
        {JointKneeRight, JointType_KneeRight},
        {JointFootRight, JointType_FootRight},
        {JointFootTipRight, JointType_FootRight},
    };

public:
    bool is_initialized()
    {
        return initialized_;
    }

    HRESULT status_result()
    {
        switch (kinect_status_result())
        {
        case S_OK: return 0;
        case S_FALSE: return 1;
        default: return -1;
        }
    }

    int initialize()
    {
        try
        {
            initialized_ = initKinect();
            if (!initialized_) return 1;

            initializeSkeleton();

            // Recreate the updater thread
            if (!updater_thread_)
                updater_thread_.reset(new std::thread(&KinectWrapper::updater, this));

            return 0; // OK
        }
        catch (...)
        {
            return -1;
        }
    }

    void update()
    {
        // Update availability
        tryCef([&, this]
        {
            if (kinectSensor)
            {
                CComPtr<IIsAvailableChangedEventArgs> args;
                if (kinectSensor->GetIsAvailableChangedEventData(h_statusChangedEvent, &args) == S_OK)
                {
                    BOOLEAN isAvailable = false;
                    args->get_IsAvailable(&isAvailable);
                    initialized_ = isAvailable; // Update the status
                    status_changed_event(); // Notify the CLR listener
                }
            }
        });

        // Update (only if the sensor is on and its status is ok)
        if (initialized_ && kinectSensor)
        {
            BOOLEAN isAvailable = false;
            HRESULT kinectStatus = kinectSensor->get_IsAvailable(&isAvailable);
            if (kinectStatus == S_OK)
            {
                // NEW ARRIVED FRAMES ------------------------
                MSG msg;
                while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) // Unnecessary?
                    DispatchMessage(&msg);

                if (h_bodyFrameEvent)
                    if (HANDLE handles[] = {reinterpret_cast<HANDLE>(h_bodyFrameEvent)};
                        // Wait for a frame to arrive, give up after >3s of nothing
                        MsgWaitForMultipleObjects(_countof(handles), handles,
                                                  false, 3000, QS_ALLINPUT) == WAIT_OBJECT_0)
                    {
                        IBodyFrameArrivedEventArgs* pArgs = nullptr;
                        if (bodyFrameReader &&
                            SUCCEEDED(bodyFrameReader->GetFrameArrivedEventData(h_bodyFrameEvent, &pArgs)))
                        {
                            [&,this](IBodyFrameReader& sender, IBodyFrameArrivedEventArgs& eventArgs)
                            {
                                updateSkeletalData();
                            }(*bodyFrameReader, *pArgs);
                            pArgs->Release(); // Release the frame
                        }
                    }
            }
        }
    }

    int shutdown()
    {
        try
        {
            // Shut down the sensor (Only NUI API)
            if (kinectSensor)
            {
                // Protect from null call
                terminateSkeleton();

                return [&, this]
                {
                    __try
                    {
                        initialized_ = false;

                        kinectSensor->Close();
                        kinectSensor->Release();

                        kinectSensor = nullptr;
                        return 0;
                    }
                    __except (EXCEPTION_EXECUTE_HANDLER)
                    {
                        return -2;
                    }
                }();
            }
            return 1;
        }
        catch (...)
        {
            return -1;
        }
    }

    void tryCef(const std::function<void()>& callback)
    {
        try
        {
            [&, this]
            {
                __try
                {
                    callback();
                }
                __except (EXCEPTION_EXECUTE_HANDLER)
                {
                }
            }();
        }
        catch (...)
        {
        }
    }

    std::array<JointOrientation, JointType_Count> bone_orientations()
    {
        return bone_orientations_;
    }

    std::array<Joint, JointType_Count> skeleton_positions()
    {
        return skeleton_positions_;
    }

    bool skeleton_tracked()
    {
        return skeleton_tracked_;
    }

    int KinectJointType(int kinectJointType)
    {
        return KinectJointTypeDictionary.at(static_cast<JointType>(kinectJointType));
    }
};
