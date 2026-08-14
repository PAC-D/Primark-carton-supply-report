from dataclasses import dataclass


@dataclass(frozen=True)
class Record:
    packaging_supplier: str
    supplier: str
    factory: str
    year: int
    month: int
    cartons: float

    @property
    def period(self):
        return self.year, self.month
