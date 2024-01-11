package main

import (
	"github.com/coalaura/logger"
)

var log = logger.New()

func main() {
	err := debugAudioSessions()
	log.MustPanic(err)
}
