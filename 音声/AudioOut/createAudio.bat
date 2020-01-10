@ECHO OFF
ffmpeg -f concat -i audio.txt -c copy output.mp3
ECHO Congratulations! Your first batch file executed successfully.
PAUSE