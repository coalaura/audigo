package main

import (
	"os"
	"os/signal"
	"syscall"
	"time"
	"unicode/utf16"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/inancgumus/screen"
	"github.com/moutend/go-wca/pkg/wca"
	"golang.org/x/sys/windows"
)

var (
	kernel32                 = syscall.NewLazyDLL("kernel32.dll")
	procGetConsoleCursorInfo = kernel32.NewProc("GetConsoleCursorInfo")
	procSetConsoleCursorInfo = kernel32.NewProc("SetConsoleCursorInfo")
)

type CONSOLE_CURSOR_INFO struct {
	Size    uint32
	Visible int32
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

	var handle syscall.Handle
	handle, err = syscall.GetStdHandle(syscall.STD_OUTPUT_HANDLE)
	if err != nil {
		return
	}

	var cci CONSOLE_CURSOR_INFO
	if err = getConsoleCursorInfo(uintptr(handle), &cci); err != nil {
		return
	}

	cci.Visible = 0
	if err = setConsoleCursorInfo(uintptr(handle), &cci); err != nil {
		return
	}

	// wait for ctrl+c signal then restore the cursor
	go func() {
		sig := make(chan os.Signal, 1)
		signal.Notify(sig, os.Interrupt, syscall.SIGTERM, syscall.SIGINT)

		<-sig

		cci.Visible = 1
		_ = setConsoleCursorInfo(uintptr(handle), &cci)

		os.Exit(0)
	}()

	for {
		screen.Clear()
		screen.MoveTopLeft()

		var sessionCount int
		if err = sessionEnumerator.GetCount(&sessionCount); err != nil {
			return
		}

		log.CPrint("- - -", 243, 0)
		log.CPrint(" Audio Sessions ", 255, 0)
		log.CPrint("- - -\n", 243, 0)

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

			var state uint32
			if err = audioSessionControl.GetState(&state); err != nil {
				continue
			}

			if state == wca.AudioSessionStateActive {
				log.CPrint(" ⬤ ", 46, 0)
			} else if state == wca.AudioSessionStateInactive {
				log.CPrint(" ⬤ ", 196, 0)
			} else {
				log.CPrint(" ⬤ ", 102, 0)
			}

			log.CPrint(processName+"\n", 248, 0)
		}

		time.Sleep(250 * time.Millisecond)
	}
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

func getConsoleCursorInfo(hConsoleOutput uintptr, cci *CONSOLE_CURSOR_INFO) error {
	ret, _, err := procGetConsoleCursorInfo.Call(
		hConsoleOutput,
		uintptr(unsafe.Pointer(cci)),
	)
	if ret == 0 {
		return err
	}

	return nil
}

func setConsoleCursorInfo(hConsoleOutput uintptr, cci *CONSOLE_CURSOR_INFO) error {
	ret, _, err := procSetConsoleCursorInfo.Call(
		hConsoleOutput,
		uintptr(unsafe.Pointer(cci)),
	)
	if ret == 0 {
		return err
	}

	return nil
}
