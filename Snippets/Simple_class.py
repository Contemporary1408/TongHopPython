class dog():
    def __init__(self, name, color):
        self.name = name
        self.color = color
    def like(self,food):
        self.food = food
        return self.name + " likes " + self.food
dog1 = dog(name = "Nick", color = "Red")
a = dog1.like(food = "pizza")
print(a)
