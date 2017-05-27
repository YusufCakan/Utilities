class Cheese(object):
    def __init__(self, num_holes=0):
        "defaults to a solid cheese"
        self.number_of_holes = num_holes

    @classmethod
    def random(cls):
        return cls(randint(0, 100))

    @classmethod
    def slightly_holey(cls):
        return cls(randint((0,33))

    @classmethod
    def very_holey(cls):
        return cls(randint(66, 100))
     
Now create object like this:

gouda = Cheese()
emmentaler = Cheese.random()
leerdammer = Cheese.slightly_holey()
