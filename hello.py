import yaml

# ---- Base class we want to inherit from ----
class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age
    
    def greet(self):
        print(f"Hello, I’m {self.name}, {self.age} years old.")

# ---- YAML-driven class generator ----
with open("classes.yaml", "r") as f:
    data = yaml.safe_load(f)

generated_classes = {}

def create_class(name, base, attributes, methods):
    # dynamic __init__ extending the base __init__
    def __init__(self, *args, **kwargs):
        # call base __init__
        super(self.__class__, self).__init__(*args, **kwargs)
        # set new attributes
        for attr in attributes:
            setattr(self, attr, kwargs.get(attr))
    
    namespace = {"__init__": __init__}
    
    for method in methods:
        def func(self, method_name=method):  # capture method name
            print(f"{method_name} called on {name}")
        namespace[method] = func
    
    return type(name, (base,), namespace)

# ---- Generate Classes ----
for class_name, props in data.items():
    base_name = props.get("base", "object")
    base_cls = globals().get(base_name, object)  # resolve base class
    attrs = props.get("attributes", [])
    methods = props.get("methods", [])
    generated_classes[class_name] = create_class(class_name, base_cls, attrs, methods)

# ---- Usage ----
Student = generated_classes["Student"]
Teacher = generated_classes["Teacher"]

s = Student(name="Alice", age=20, roll_no="101", grade="A")
t = Teacher(name="Bob", age=40, subject="Math")

s.greet()
s.study()
print(s.roll_no, s.grade)

t.greet()
t.teach()
print(t.subject)
