def smart_divide(func):
    def inner(a,b):
        print("I am going to divide",a,"and",b)
        if b == 0:
            print("Whoops! cannot divide")
            return  
        print(func(a,b))
    return inner
  
@smart_divide
def divide(a,b):
    return a / b
divide(a = 8,b = 4)
