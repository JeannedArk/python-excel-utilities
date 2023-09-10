this_is_an_integer = 4545
this_is_a_string = "hello"
this_is_an_array_of_integers = [1, 2, 3, 4]
this_is_an_array_of_strings = ["1", "wobbly", "3"]

# def print_a_string(param: str):
def print_a_string(param):
    print("String: " + param)

print_a_string(this_is_a_string)

# def print_strings(params: list[str]):
def print_strings(params):
    for param in params:
        print("like string column" + param)

print_strings(this_is_an_array_of_strings)

def print_integer(param: int):
    print("integer" + str(param))

print_integer(this_is_an_integer)

def print_integer_array(params):
    for param in params:
        print("integer array" + str(param))

print_integer_array(this_is_an_array_of_integers)