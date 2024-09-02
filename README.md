<p align="center"> <img width="256" height="256" src="https://github.com/jame25/Piper-Tray/assets/13631646/95cf62c7-9248-4195-864b-57771d922bd6"></p>


Piper Tray (Classic) is a small system tray utility for Windows that utilizes [Piper](https://github.com/rhasspy/piper). It will read aloud the contents of your clipboard. You can stop the speech at any time either by using the 'Stop Speech' option, or via the hotkey (Alt + Q).

If you want more features, see: [Piper Tray](https://github.com/jame25/Piper-Tray). Alternatively, if you prefer a GUI; check out my other project: [Piper Read](https://github.com/jame25/piper-read).

## Features:

* Reads clipboard contents aloud
* Enable / Disable clipboard monitoring
* Many voices to choose from
* Change Piper TTS voice model
* Control Piper TTS speech rate
* Hotkey support (ALT + Q) to stop speech
* Prevent keywords being read (ignore.dict)
* Replace keywords with alternatives (replace.dict)

## Prerequisites:

[.Net 8.0 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) is required to be installed.

## Install:

- Download the latest version of Piper Tray from [releases](https://github.com/jame25/Piper-Tray-Classic/releases/).
- Grab the latest Windows binary for Piper from [here](https://github.com/rhasspy/piper/releases). Voice models are available [here](https://huggingface.co/rhasspy/piper-voices/tree/main).
- <b>Extract all of the above into the same directory</b>.

## Configuration:

Piper Tray should support all available Piper voice models, by default **en_US-libritts_r-medium.onnx** and .json are expected to be present in directory.

## Settings:

You can change the voice model being utilized by Piper Tray by editing the first line of the **settings.conf**.

Speech rate can be altered using the 'speed' variable (1.0 is the default speed, lower value i.e 0.5 = faster).

## Dictionary Rules:

Keywords found in the **ignore.dict** file are skipped over. 

If a keyword in the **banned.dict** file is detected, the entire line is skipped.

**replace.dict** functions as a replacement for a keyword or phrase, i.e LHC=Large Hadron Collider

## Support:

If you find this project helpful and would like to support its development, you can buy me a coffee on Ko-Fi:

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/jame25)
