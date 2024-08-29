import wave
import pyaudio

def play_stop():
    chunk = 1024
    wf = wave.open('C:\\Users\\Alexeenko\\Travian\\new_message-6.wav', 'rb')
    pa = pyaudio.PyAudio()

    stream = pa.open(
        format=pa.get_format_from_width(wf.getsampwidth()),
        channels=wf.getnchannels(),
        rate=wf.getframerate(),
        output=True
    )

    data = wf.readframes(chunk)

    while data:
        stream.write(data)
        data = wf.readframes(chunk)

    stream.close()
    pa.terminate()

def play_build():
    chunk = 1024
    wf = wave.open('C:\\Users\\Alexeenko\\Travian\\build.wav', 'rb')
    pa = pyaudio.PyAudio()

    stream = pa.open(
        format=pa.get_format_from_width(wf.getsampwidth()),
        channels=wf.getnchannels(),
        rate=wf.getframerate(),
        output=True
    )

    data = wf.readframes(chunk)

    while data:
        stream.write(data)
        data = wf.readframes(chunk)

    stream.close()
    pa.terminate()
