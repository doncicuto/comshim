package comshim_test

import (
	"log"

	"github.com/doncicuto/comshim"
)

func Example_globalUsage() {
	// This ensures that at least one thread maintains an initialized
	// multi-threaded COM apartment.
	if err := comshim.Add(1); err != nil {
		log.Printf("[ERROR]: could not add thread %v", err)
		return
	}

	// After we're done using COM the thread will be released.
	defer func() {
		if err := comshim.Done(); err != nil {
			log.Printf("[ERROR]: Done call found an error %v", err)
			return
		}
	}()

	// Do COM things here
}
