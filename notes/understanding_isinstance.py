"""
What is isinstance()?

isinstance checks if an object is a specific type

Those types can be as easy as Integers, Floats, and Strings
or more complex like Class and Subclass.
"""

# Checks to see if 5 is an integer
if isinstance(5, int):
    print('isinstance(5, int) is True')
else:
    print('isinstance(5, int) is False')

# Checks to see if 'hello' is a string
if isinstance('hello', str):
    print('isinstance("hello", str) is True')
else:
    print('isinstance("hello", str) is False')

# Checks to see if 5 is a string
if isinstance(5, str):
    print('isinstance(5, str) is True')
else:
    print('isinstance(5, str) is False')

# True and False are technically 1 and 0 under the hood and will trip Integer check
if isinstance(True, int):
    print('isinstance(True, int) is True')

if isinstance(True, bool):
    print('isinstance(True, bool) is True')

if type(True) == int:
    print('(type(True) == int) is True')
else:
    print('(type(True) == int) is False')

if isinstance(True, int):
    print('isinstance(True, int) is True')
else:
    print('isinstance(True, int) is False')

# Will need to guard against it sometimes
number = True
if isinstance(number, int) and not isinstance(number, bool):
    print(f'isinstance({number}, int) and not isinstance({number}, bool) is True')
else:
    print(f'isinstance({number}, int) and not isinstance({number}, bool) is False')

# Checks to see if d is an int or Float
if isinstance(5, (int,float)):
    print('isinstance(5, (int, float)) is True')
else:
    print('isinstance(5, (int, float)) is False')

class Animal:
    def __init__(self, species:str, gender:str):
        self.species = species
        self.gender = gender

doggo = Animal('dog', 'male' )

# Checks to see if doggo is an animal
if isinstance(doggo, Animal):
    print('isinstance(doggo, Animal) is True')
else:
    print('isinstance(doggo, Animal) is False')

# Avoid doing this
x = 5
print(type(x)== int)
print(type(x)== float)

# Use this
print(isinstance(x, int))
