# This is a function (definition)

def noparameters():
    print("no parameters")

def oneparameter(parameter1):
    print("This is parameter: " + parameter1)

def twoparameters(parameter1, parameter2):
    print("has two parameters: " + parameter1 + " " + parameter2)

def fiveparameters(pm1, pm2, pm3, pm4, pm5):
    print("has five parameters" + pm1, pm2, pm3, pm4, pm5)

# Syntax of calling a function:
# function_name(parameters...)

# Calling built-in print function
#print("hello laura")
# Calling your function:
#toprintparameter(1)

noparameters()
oneparameter("This makes sense")
twoparameters("string 1", "s2")
fiveparameters("s1", "s2", "s3", "s4", "s5")