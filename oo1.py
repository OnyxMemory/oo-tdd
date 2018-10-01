
class Person:

    def __init__(self,first, last, age):
        self.first_name = first
        self.last_name = last
        self.age = age

    @property
    def name(self):
        return f'{self.first_name} {self.last_name}'

    def birthday(self):
        self.age += 1



