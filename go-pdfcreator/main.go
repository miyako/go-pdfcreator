package main

import (
	"flag"
	"fmt"
	"os"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {

	wait := flag.Int("wait", 60, "wait")
	path := flag.String("path", "", "path")

	flag.Parse()
	if *path == "" {
		fmt.Fprintf(os.Stderr, "path can't be empty!")
		os.Exit(1)
	}

	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("PDFCreator.JobQueue")
	if err != nil {
		fmt.Fprintln(os.Stderr, "Failed to create COM object:", err)
		os.Exit(1)
	}

	queue, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		fmt.Fprintln(os.Stderr, "Failed to get IDispatch:", err)
		os.Exit(1)
	}
	defer queue.Release()

	_, err = oleutil.CallMethod(queue, "Initialize")
	if err != nil {
		fmt.Fprintln(os.Stderr, "Failed to initialize queue:", err)
		os.Exit(1)
	}
	defer oleutil.CallMethod(queue, "ReleaseCom")

	fmt.Fprintln(os.Stderr, "Waiting for a job to be sent to PDFCreator printer...")

	result, err := oleutil.CallMethod(queue, "WaitForJob", wait)
	if err != nil || result.Val == 0 {
		fmt.Fprintln(os.Stderr, "No job received within timeout.")
		os.Exit(1)
	}

	jobResult, err := oleutil.CallMethod(queue, "NextJob")
	if err != nil {
		fmt.Fprintln(os.Stderr, "Failed to get next job:", err)
		os.Exit(1)
	}
	job := jobResult.ToIDispatch()
	defer job.Release()

	oleutil.CallMethod(job, "SetProfileByGuid", "DefaultGuid")

	fmt.Fprintln(os.Stderr, "Converting print job to PDF:", *path)
	_, err = oleutil.CallMethod(job, "ConvertTo", *path)
	if err != nil {
		fmt.Fprintln(os.Stderr, "Conversion failed:", err)
		os.Exit(1)
	}

	for {
		result, _ := oleutil.CallMethod(job, "IsFinished")
		if result.Val != 0 {
			break
		}

		progress, _ := oleutil.CallMethod(job, "GetProgress")
		fmt.Fprintln(os.Stderr, "\rProgress: %d%%", progress.Val)

		time.Sleep(500 * time.Millisecond)
	}

	// Check if conversion succeeded
	success, err := oleutil.CallMethod(job, "IsSuccessful")
	if err != nil {
		fmt.Fprintln(os.Stderr, "Error checking success:", err)
		os.Exit(1)
	}

	if success.Val != 0 {
		fmt.Fprintln(os.Stderr, "PDF conversion succeeded.")
	} else {
		fmt.Fprintln(os.Stderr, "PDF conversion failed.")
		os.Exit(1)
	}
}
