from dataclasses import dataclass


@dataclass
class Element:
    kind: str
    x: float
    y: float
    w: float
    h: float
    data: dict
