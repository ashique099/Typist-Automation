"""Per-job API usage + cost, and per-job model/effort settings (thread-local, one job per thread)."""
from __future__ import annotations

import threading

# USD per 1M tokens: input, output, cache read, cache write (5 min)
PRICES = {
    "claude-opus-5-5": (4.0, 20.0, 0.20, 5.0),
    "claude-sonnet-5-5": (2.0, 10.0, 0.10, 2.5),
    "claude-haiku-5-5": (0.10, 0.50, 0.01, 0.125),
}
MODELS = {"claude-opus-5-5": "Opus 5.5 (best quality)", "claude-sonnet-5-5": "Sonnet 5.5 (aadha kharcha)",
          "claude-haiku-5-5": "Haiku 5.5 (sabse sasta)"}

_local = threading.local()


class Meter:
    def __init__(self, model=None, effort=None):
        self.model, self.effort = model, effort
        self.calls = self.input = self.output = 0
        self.cost = 0.0

    def add(self, model, usage):
        pin, pout, pcr, pcw = PRICES.get(model, PRICES["claude-opus-5-5"])
        i = getattr(usage, "input_tokens", 0) or 0
        o = getattr(usage, "output_tokens", 0) or 0
        cr = getattr(usage, "cache_read_input_tokens", 0) or 0
        cw = getattr(usage, "cache_creation_input_tokens", 0) or 0
        self.calls += 1
        self.input += i + cr + cw
        self.output += o
        self.cost += (i * pin + o * pout + cr * pcr + cw * pcw) / 1e6

    def as_dict(self):
        return {"calls": self.calls, "input_tokens": self.input, "output_tokens": self.output,
                "cost_usd": round(self.cost, 4)}


def start(model=None, effort=None) -> Meter:
    _local.meter = Meter(model, effort)
    return _local.meter


def current():
    return getattr(_local, "meter", None)


def model(default):
    m = current()
    return (m and m.model) or default


def effort(default):
    m = current()
    return (m and m.effort) or default


def record(msg):
    m = current()
    if m is not None and getattr(msg, "usage", None) is not None:
        m.add(getattr(msg, "model", None) or m.model, msg.usage)
