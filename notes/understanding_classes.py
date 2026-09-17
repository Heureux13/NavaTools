"""
A class is a blueprint of an object, when something is created with a class it is called an Instance
"""


class Car:
    def __init__(self, make: str, model: str, year: int):
        self.make = make
        self.model = model
        self.year = year

    def start(self) -> str:
        return f'{self.make} {self.model} Car starts'

    def info(self) -> str:
        return f'{self.year} {self.make} {self.model}'

    def __repr__(self) -> str:
        return self.info()


# car00 and car01 are both objects of the Car class but more specifically they are Instances of Car class
car00 = Car('Toyota', "T100", 1996)
car01 = Car("Chevy", "S-10", 1991)

print(car00.info())
print(car01.info())
print(car00)
