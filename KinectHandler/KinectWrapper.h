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
    IMultiSourceFrameReader* multiFrameReader = nullptr;
    IMultiSourceFrame* multiFrame = nullptr;
    ICoordinateMapper* coordMapper = nullptr;
    BOOLEAN isTracking = false;

    Joint joints[JointType_Count];
    JointOrientation boneOrientations[JointType_Count];
    IBody* kinectBodies[BODY_COUNT] = {nullptr};

    WAITABLE_HANDLE h_statusChangedEvent;
    WAITABLE_HANDLE h_multiFrameEvent;

    std::array<JointOrientation, JointType_Count> bone_orientations_;
    std::array<Joint, JointType_Count> skeleton_positions_;

    std::unique_ptr<std::thread> updater_thread_;

    inline static bool initialized_ = false;
    bool skeleton_tracked_ = false;
    bool rgb_stream_enabled_ = false;

    void updater()
    {
        // Auto-handles failures & etc
        while (true)update();
    }

    void updateFrameData(IMultiSourceFrameArrivedEventArgs* args)
    {
        if (!multiFrameReader) return; // Give up already

        // Acquire the multi-source frame reference
        IMultiSourceFrameReference* frameReference = nullptr;
        args->get_FrameReference(&frameReference);

        IMultiSourceFrame* multiFrame = nullptr;
        frameReference->AcquireFrame(&multiFrame);

        if (!multiFrame) return;

        // Get the body frame and process it
        IBodyFrameReference* bodyFrameReference = nullptr;
        multiFrame->get_BodyFrameReference(&bodyFrameReference);

        IBodyFrame* bodyFrame = nullptr;
        bodyFrameReference->AcquireFrame(&bodyFrame);

        if (!bodyFrame) return;
        bodyFrame->GetAndRefreshBodyData(BODY_COUNT, kinectBodies);
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

        // Don't process color if not requested
        if (!camera_enabled()) return;
        
        // Get the color frame and process it
        IColorFrameReference* colorFrameReference = nullptr;
        multiFrame->get_ColorFrameReference(&colorFrameReference);

        IColorFrame* colorFrame = nullptr;
        colorFrameReference->AcquireFrame(&colorFrame);

        if (!colorFrame) return;
        ResetBuffer(CameraBufferSize()); // Allocate buffer for image for copy

        colorFrame->CopyConvertedFrameDataToArray(
            CameraBufferSize(), color_buffer_, ColorImageFormat_Bgra);

        if (colorFrame) colorFrame->Release();
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

#ifdef _DEBUG
            // Emulation support bypass
            available = true;
#endif

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

    void initializeFrameReader()
    {
        if (multiFrameReader)
            multiFrameReader->Release();

        kinectSensor->OpenMultiSourceFrameReader(
            FrameSourceTypes_Body | FrameSourceTypes_Color, &multiFrameReader);

        // Newfangled event based frame capture
        // https://github.com/StevenHickson/PCL_Kinect2SDK/blob/master/src/Microsoft_grabber2.cpp
        h_multiFrameEvent = (WAITABLE_HANDLE)CreateEvent(nullptr, FALSE, FALSE, nullptr);
        multiFrameReader->SubscribeMultiSourceFrameArrived(&h_multiFrameEvent);
    }

    void terminateMultiFrame()
    {
        if (!multiFrameReader)return; // No need to do anything
        if (FAILED(multiFrameReader->UnsubscribeMultiSourceFrameArrived(h_multiFrameEvent)))
        {
            throw std::exception("Couldn't unsubscribe frame!");
        }
        __try
        {
            CloseHandle((HANDLE)h_multiFrameEvent);
            multiFrameReader->Release();
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            // ignored
        }
        h_multiFrameEvent = NULL;
        multiFrameReader = nullptr;
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
        {JointFootLeft, JointType_AnkleLeft},
        {JointFootTipLeft, JointType_FootLeft},
        {JointHipRight, JointType_HipRight},
        {JointKneeRight, JointType_KneeRight},
        {JointFootRight, JointType_AnkleRight},
        {JointFootTipRight, JointType_FootRight},
    };

public:
    bool is_initialized()
    {
#ifdef _DEBUG
        return true;
#else
        return initialized_;
#endif
    }

    HRESULT status_result()
    {
#ifdef _DEBUG
        return 0;
#else
        switch (kinect_status_result())
        {
        case S_OK: return 0;
        case S_FALSE: return 1;
        default: return -1;
        }
#endif
    }

    int initialize()
    {
        try
        {
            initialized_ = initKinect();
            if (!is_initialized()) return 1;

            initializeFrameReader();

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

#ifdef _DEBUG
                    // Emulation support bypass
                    isAvailable = true;
#endif

                    initialized_ = isAvailable; // Update the status
                    status_changed_event(); // Notify the CLR listener
                }
            }
        });

        // Update (only if the sensor is on and its status is ok)
        if (is_initialized() && kinectSensor)
        {
            BOOLEAN isAvailable = false;
            HRESULT kinectStatus = kinectSensor->get_IsAvailable(&isAvailable);

#ifdef _DEBUG
            // Emulation support bypass
            isAvailable = true;
#endif

            if (kinectStatus == S_OK)
            {
                // NEW ARRIVED FRAMES ------------------------
                MSG msg;
                while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) // Unnecessary?
                    DispatchMessage(&msg);

                if (h_multiFrameEvent)
                    if (HANDLE handles[] = {reinterpret_cast<HANDLE>(h_multiFrameEvent)};
                        // Wait for a frame to arrive, give up after >3s of nothing
                        MsgWaitForMultipleObjects(_countof(handles), handles,
                                                  false, 100, QS_ALLINPUT) == WAIT_OBJECT_0)
                    {
                        IMultiSourceFrameArrivedEventArgs* pArgs = nullptr;
                        if (multiFrameReader &&
                            SUCCEEDED(multiFrameReader->GetMultiSourceFrameArrivedEventData(h_multiFrameEvent, &pArgs)))
                        {
                            [&,this](IMultiSourceFrameReader& sender, IMultiSourceFrameArrivedEventArgs& eventArgs)
                            {
                                updateFrameData(&eventArgs);
                            }(*multiFrameReader, *pArgs);
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
                terminateMultiFrame();

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

    std::tuple<BYTE*, int> color_buffer()
    {
        return std::make_tuple(color_buffer_, size_in_bytes_last_);
    }

    bool skeleton_tracked()
    {
        return skeleton_tracked_;
    }

    void camera_enabled(bool enabled)
    {
        rgb_stream_enabled_ = enabled;
    }

    bool camera_enabled(void)
    {
        return rgb_stream_enabled_;
    }

    int KinectJointType(int kinectJointType)
    {
        return KinectJointTypeDictionary.at(static_cast<JointType>(kinectJointType));
    }

    std::pair<int, int> CameraImageSize()
    {
        return std::make_pair(1920, 1080);
    }

    unsigned long CameraBufferSize()
    {
        const auto& [width, height] = CameraImageSize();
        return width * height * 4;
    }

    std::pair<int, int> MapCoordinate(const CameraSpacePoint& skeletonPoint)
    {
        auto _skeletonPoint = skeletonPoint;
        if (_skeletonPoint.Z < 0) _skeletonPoint.Z = 0.1f;

        ColorSpacePoint spacePoint; // Holds the mapped coordinate
        const auto& result = coordMapper->MapCameraPointToColorSpace(_skeletonPoint, &spacePoint);

        return SUCCEEDED(result) && !std::isnan(spacePoint.X) && !std::isnan(spacePoint.Y)
                   ? std::make_pair(spacePoint.X, spacePoint.Y) // Send the mapped ones
                   : std::make_pair(-1, -1); // Unknown coordinates - fall back to default drawing
    }

private:
    DWORD size_in_bytes_ = 0;
    DWORD size_in_bytes_last_ = 0;
    BYTE* color_buffer_ = nullptr;

    BYTE* ResetBuffer(UINT size)
    {
        size_in_bytes_last_ = size;
        if (!color_buffer_ || size_in_bytes_ != size)
        {
            if (color_buffer_)
            {
                delete[] color_buffer_;
                color_buffer_ = nullptr;
            }

            if (0 != size)
            {
                color_buffer_ = new BYTE[size];
            }
            size_in_bytes_ = size;
        }

        return color_buffer_;
    }
};
