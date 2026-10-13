"""Synthesizes the 10 s loopable nursery-rhyme soundtrack + sound effects -> soundtrack.wav.

120 BPM (1 beat = 0.5 s), key of C. Events line up with scene.js:
  0.00-0.75 egg wobble (tick-tock)   0.75 POP + confetti sparkles   1.15 giggle
  2/4/6 s   dance reps; "boing" on each hop (beat 3), balloon "pop" on beat 4
  8.0       balloons float up (whistle) + goodbye   9.3-10 egg closes (loops)
"""
import wave
import numpy as np

SR = 44100
LEN = 10.0
BEAT = 0.5
rng = np.random.default_rng(3)
buf = np.zeros(int(SR * (LEN + 2.5)))  # extra tail gets wrapped to the start for a seamless loop


def add(sig, t, gain=1.0):
    i = int(round(t * SR))
    buf[i:i + len(sig)] += sig[: len(buf) - i] * gain


def tt(d):
    return np.arange(int(d * SR)) / SR


def hz(name):
    names = {'C': 0, 'D': 2, 'E': 4, 'F': 5, 'G': 7, 'A': 9, 'B': 11}
    n = names[name[0]] + (1 if '#' in name else 0)
    octave = int(name[-1])
    return 440.0 * 2 ** ((n + 12 * (octave + 1) - 69) / 12)


# ---------------- instruments ----------------
def mallet(f, d=0.9):
    """Glockenspiel/marimba-like melody voice."""
    t = tt(d)
    att = 1 - np.exp(-t / 0.003)
    s = (np.sin(2 * np.pi * f * t) * np.exp(-t / 0.45)
         + 0.35 * np.sin(2 * np.pi * 2 * f * t) * np.exp(-t / 0.2)
         + 0.12 * np.sin(2 * np.pi * 3.98 * f * t) * np.exp(-t / 0.06))
    return s * att


def bass(f, d=0.45):
    t = tt(d)
    s = np.sin(2 * np.pi * f * t) + 0.25 * np.sin(2 * np.pi * 2 * f * t)
    return s * (1 - np.exp(-t / 0.004)) * np.exp(-t / 0.22)


def uke(freqs, d=0.6):
    """Karplus-Strong strum (ukulele-ish), strings slightly staggered."""
    out = np.zeros(int(d * SR) + 600)
    for k, f in enumerate(freqs):
        n = int(SR / f)
        line = rng.uniform(-1, 1, n)
        y = np.zeros(int(d * SR))
        for i in range(len(y)):
            y[i] = line[i % n]
            line[i % n] = 0.497 * (line[i % n] + line[(i + 1) % n])
        o = k * 120
        out[o:o + len(y)] += y
    return out


def kick():
    t = tt(0.25)
    f = 50 + 90 * np.exp(-t / 0.03)
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t / 0.09)


def shaker():
    t = tt(0.07)
    n = np.diff(rng.uniform(-1, 1, len(t) + 1))
    return n * np.exp(-t / 0.018)


# ---------------- sound effects ----------------
def tock(f):
    t = tt(0.12)
    return np.sin(2 * np.pi * f * t) * np.exp(-t / 0.03)


def pop():
    t = tt(0.25)
    f = 900 * np.exp(-t / 0.03) + 140
    body = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t / 0.06)
    noise = rng.uniform(-1, 1, len(t)) * np.exp(-t / 0.012)
    return body + 0.6 * noise


def cork():
    """Little bright 'bloop' for a balloon appearing."""
    t = tt(0.16)
    f = 350 + 1100 * (1 - np.exp(-t / 0.03))
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * (1 - np.exp(-t / 0.003)) * np.exp(-t / 0.05)


def boing():
    t = tt(0.55)
    f = 160 + 280 * (1 - np.exp(-t / 0.12)) + 40 * np.sin(2 * np.pi * 16 * t) * np.exp(-t / 0.2)
    ph = 2 * np.pi * np.cumsum(f) / SR
    return (np.sin(ph) + 0.3 * np.sin(2 * ph)) * (1 - np.exp(-t / 0.005)) * np.exp(-t / 0.18)


def giggle(start_f, n=5, gap=0.11):
    """'hee-hee-hee': harmonic voice shaped by an 'ee' vowel formant envelope."""
    out = np.zeros(int((n * gap + 0.2) * SR))
    for k in range(n):
        f0 = start_f * (1 + 0.06 * np.sin(k * 1.7)) * (1 - 0.03 * k)
        t = tt(0.1)
        f = f0 * (1 + 0.05 * np.sin(2 * np.pi * 9 * t))
        ph = 2 * np.pi * np.cumsum(f) / SR
        s = np.zeros_like(t)
        for h in range(1, 12):
            fh = f0 * h
            w = np.exp(-((fh - 320) / 180) ** 2) + 0.7 * np.exp(-((fh - 2600) / 500) ** 2) + 0.05
            s += w * np.sin(h * ph) / h ** 0.3
        env = np.minimum(t / 0.012, 1) * np.exp(-t / 0.045)
        i = int(k * gap * SR)
        out[i:i + len(t)] += s * env
    return out / np.max(np.abs(out))


def whistle():
    t = tt(1.0)
    f = 500 + 900 * t ** 1.3 + 15 * np.sin(2 * np.pi * 6 * t)
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * np.minimum(t / 0.1, 1) * np.exp(-t / 0.5)


def sparkles(t0, dur, count):
    for _ in range(count):
        f = rng.uniform(2000, 4200)
        t = tt(0.3)
        add(np.sin(2 * np.pi * f * t) * np.exp(-t / 0.07), t0 + rng.uniform(0, dur), 0.05)


# ---------------- music ----------------
CHORDS = {'C': ['C', 'E', 'G'], 'F': ['F', 'A', 'C'], 'G': ['G', 'B', 'D']}
bars = ['C', 'F', 'G', 'C']  # bars 1-4 (t = 2..10); bar 0 is the egg intro

# intro: rising "tick-tock" wobble notes, then a glissando sparkle on the pop
for t, n in [(0.0, 'G4'), (0.25, 'A4'), (0.5, 'B4')]:
    add(mallet(hz(n), 0.4), t, 0.35)
for i, n in enumerate(['C5', 'E5', 'G5', 'C6', 'E6']):
    add(mallet(hz(n)), 0.75 + i * 0.05, 0.35)
add(bass(hz('C3'), 0.8), 0.75, 0.3)
add(uke([hz('C4'), hz('E4'), hz('G4'), hz('C5')], 1.0), 1.0, 0.25)
for t, n in [(1.5, 'G4'), (1.75, 'B4')]:  # pick-up into the dance
    add(mallet(hz(n), 0.4), t, 0.3)

# the catchy repeating motif, one per bar (last note lands with each balloon)
melody = {
    'C': ['E5', 'G5', 'E5', 'G5', 'A5', 'G5', 'C6'],
    'F': ['F5', 'A5', 'F5', 'A5', 'C6', 'A5', 'F5'],
    'G': ['G5', 'B5', 'G5', 'B5', 'D6', 'B5', 'G5'],
}
rhythm = [0, 0.5, 1, 1.5, 2, 2.5, 3]  # beats
goodbye = [(0, 'G5'), (0.5, 'E5'), (1, 'G5'), (1.5, 'E5'), (2, 'C5'), (3, 'G4'), (3.5, 'C5')]

for b, ch in enumerate(bars):
    t0 = 2 + b * 4 * BEAT
    notes = list(zip(rhythm, melody[ch])) if b < 3 else goodbye
    for beat, n in notes:
        add(mallet(hz(n)), t0 + beat * BEAT, 0.42 if beat < 3 else 0.5)
    root = hz(ch + '2') if ch != 'C' else hz('C3')
    fifth = root * 1.5
    for beat, f in [(0, root), (1, fifth), (2, root), (3, fifth)]:
        add(bass(f), t0 + beat * BEAT, 0.28)
        add(kick(), t0 + beat * BEAT, 0.2 if beat % 2 == 0 else 0.08)
    tri = [hz(n + '4') for n in CHORDS[ch]]
    for beat in (0.5, 1.5, 2.5, 3.5):  # off-beat ukulele "chucks"
        add(uke(tri, 0.45), t0 + beat * BEAT, 0.16)
    for e in range(8):
        add(shaker(), t0 + e * 0.25, 0.05 if e % 2 else 0.03)

# ---------------- sound effects on cue ----------------
for t, f in [(0.05, 820), (0.2, 640), (0.35, 820), (0.5, 640), (0.62, 820)]:
    add(tock(f), t, 0.2)                       # egg wobble
add(pop(), 0.75, 0.75)                         # egg POP
sparkles(0.75, 0.8, 26)                        # confetti twinkles
add(giggle(620), 1.15, 0.3)                    # baby dino giggle
for T0 in (2, 4, 6):
    add(boing(), T0 + 1.0, 0.45)               # hop
    add(cork(), T0 + 1.5, 0.4)                 # balloon appears
add(whistle(), 8.0, 0.22)                      # balloons float away
add(giggle(700, n=3), 8.55, 0.22)              # happy goodbye giggle
add(tock(900), 9.4, 0.2)                       # egg bottom returns
add(tock(1100), 9.85, 0.25)                    # egg snaps shut

# ---------------- mix + seamless loop ----------------
n = int(LEN * SR)
tail = buf[n:]
out = buf[:n].copy()
out[:len(tail)] += tail                        # wrap ringing tails around to the loop start
out = np.tanh(out * 1.2)                       # gentle limiter
out *= 0.89 / np.max(np.abs(out))
stereo = np.stack([out, out], axis=1)
pcm = (stereo * 32767).astype('<i2')
with wave.open('soundtrack.wav', 'wb') as w:
    w.setnchannels(2)
    w.setsampwidth(2)
    w.setframerate(SR)
    w.writeframes(pcm.tobytes())
print('wrote soundtrack.wav', len(out) / SR, 's')
