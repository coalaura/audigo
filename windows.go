package main

import (
	"fmt"
	"time"
	"unicode/utf16"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/moutend/go-wca/pkg/wca"
	"golang.org/x/sys/windows"
)

type AudioSessionInfo struct {
	ProcessName    string
	SessionControl *wca.IAudioSessionControl
}

func debugAudioSessions() (err error) {
	log.Debug("Initializing Windows Core Audio API")

	if err = ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); err != nil {
		return
	}
	defer ole.CoUninitialize()

	var deviceEnumerator *wca.IMMDeviceEnumerator
	if err = wca.CoCreateInstance(wca.CLSID_MMDeviceEnumerator, 0, wca.CLSCTX_ALL, wca.IID_IMMDeviceEnumerator, &deviceEnumerator); err != nil {
		return
	}
	defer deviceEnumerator.Release()

	var defaultDevice *wca.IMMDevice
	if err = deviceEnumerator.GetDefaultAudioEndpoint(wca.ERender, wca.EConsole, &defaultDevice); err != nil {
		return
	}
	defer defaultDevice.Release()

	var audioSessionManager *wca.IAudioSessionManager2
	if err = defaultDevice.Activate(wca.IID_IAudioSessionManager2, wca.CLSCTX_ALL, nil, &audioSessionManager); err != nil {
		return
	}

	var sessionEnumerator *wca.IAudioSessionEnumerator
	if err = audioSessionManager.GetSessionEnumerator(&sessionEnumerator); err != nil {
		return
	}
	defer sessionEnumerator.Release()

	var sessionCount int
	if err = sessionEnumerator.GetCount(&sessionCount); err != nil {
		return
	}

	var sessions []AudioSessionInfo

	for i := 0; i < sessionCount; i++ {
		var audioSessionControl *wca.IAudioSessionControl
		if err := sessionEnumerator.GetSession(i, &audioSessionControl); err != nil {
			continue
		}
		defer audioSessionControl.Release()

		var dispatch *ole.IDispatch
		dispatch, err = audioSessionControl.QueryInterface(wca.IID_IAudioSessionControl2)
		if err != nil {
			continue
		}

		audioSessionControl2 := (*wca.IAudioSessionControl2)(unsafe.Pointer(dispatch))
		defer audioSessionControl2.Release()

		var processId uint32
		if err := audioSessionControl2.GetProcessId(&processId); err != nil {
			continue
		}

		processName, err := getProcessName(processId)
		if err != nil {
			continue
		}

		sessions = append(sessions, AudioSessionInfo{
			ProcessName:    processName,
			SessionControl: audioSessionControl,
		})
	}

	if len(sessions) == 0 {
		err = fmt.Errorf("no audio sessions found")

		return
	}

	for index, session := range sessions {
		log.InfoF("[%d] %s\n", index, session.ProcessName)
	}

	var index int

	log.InfoF("> ")
	fmt.Scanf("%d", &index)

	if index < 0 || len(sessions) <= index {
		err = fmt.Errorf("invalid session")

		return
	}

	session := sessions[index]

	log.InfoF("Monitoring %s\n", session.ProcessName)

	var (
		previousState uint32
		state         uint32
	)

	for {
		if err = session.SessionControl.GetState(&state); err != nil {
			return
		}

		if state == wca.AudioSessionStateExpired {
			log.Warning("Audio session has expired")

			break
		}

		if previousState != state {
			log.NoteF("%s -> %s\n", stateToString(previousState), stateToString(state))
		}

		previousState = state

		time.Sleep(200 * time.Millisecond)
	}

	return
}

func getProcessName(pid uint32) (string, error) {
	hProcess, err := windows.OpenProcess(windows.PROCESS_QUERY_INFORMATION|windows.PROCESS_VM_READ, false, pid)
	if err != nil {
		return "", err
	}
	defer windows.CloseHandle(hProcess)

	var processName [1024]uint16
	err = windows.GetModuleBaseName(hProcess, 0, &processName[0], 1024)
	if err != nil {
		return "", err
	}

	i := 0
	for i < 1024 && processName[i] != 0 {
		i++
	}

	return string(utf16.Decode(processName[:i])), nil
}

func stateToString(state uint32) string {
	switch state {
	case wca.AudioSessionStateInactive:
		return "Inactive"
	case wca.AudioSessionStateActive:
		return "Active"
	case wca.AudioSessionStateExpired:
		return "Expired"
	default:
		return "Unknown"
	}
}
