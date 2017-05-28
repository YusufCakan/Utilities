### Working with the Python Super Function
Python 2.2 saw the introduction of a built-in function called “super,” which returns a proxy object to delegate method calls to a class – which can be either parent or sibling in nature.

That description may not make sense unless you have experience working with Python, so we’ll break it down.

Essentially, the super function can be used to gain access to inherited methods – from a parent or sibling class – that has been overwritten in a class object.

Or, as the official Python documentation says:


*“Super is used to] return a proxy object that delegates method calls to a parent or sibling class of type. This is useful for accessing inherited methods that have been overridden in a class. The search order is same as that used by getattr() except that the type itself is skipped.”*


#### How Is the Super Function Used?
The super function is somewhat versatile, and can be used in a couple of ways.

**Use Case 1:** Super can be called upon in a single inheritance, in order to refer to the parent class or multiple classes without explicitly naming them. It’s somewhat of a shortcut, but more importantly, it helps keep your code maintainable for the foreseeable future.

**Use Case 2:** Super can be called upon in a dynamic execution environment for multiple or collaborative inheritance. This use is considered exclusive to Python, because it’s not possible with languages that only support single inheritance or are statically compiled.

When the super function was introduced it sparked a bit of controversy. Many developers found the documentation unclear, and the function itself to be tricky to implement. It even garnered a reputation for being harmful. But it’s important to remember that Python has evolved considerably since 2.2 and many of these concerns no longer apply.

The great thing about super is that it can be used to enhance any module method. Plus, there’s no need to know the details about the base class that’s being used as an extender. The super function handles all of it for you.

So, for all intents and purposes, super is a shortcut to access a base class without having to know its type or name.

In Python 3 and above, the syntax for super is:

     super().methoName(args)

Whereas the normal way to call super (in older builds of Python) is:


    super(subClass, instance).method(args)

As you can see, the newer version of Python makes the syntax a little simpler.



How to Call Super in Python 2 and Python 3?
First, we’ll take a regular class definition and modify it by adding the super function. The initial code will look something like this:

 

     class MyParentClass(object):
          def __init__(self):
               pass
          
     class SubClass(MyParentClass):
          def __init__(self):
               MyParentClass.__init__(self)
 

As you can see, this is a setup commonly used for single inheritance. We can see that there’s a base or parent class (also sometimes called the super class), and a denoted subclass.

But we still need to initialize the parent class within the subclass. To make this process easier, Python’s core development team created the super function. The goal was to provide a much more abstract and portable solution for initializing classes.

If we were using Python 2, we would write the subclass like this (using the super function):

 

     class SubClass(MyParentClass):
          def __init__(self):
               super(SubClass, self).__init__()
 

The same code is slightly different when writing in Python 3, however.

     class MyParentClass():
          def __init__(self):
               pass

     class SubClass(MyParentClass):
          def __init__(self):
               super()
 

Notice how the parent class isn’t directly based on the object base class anymore? In addition, thanks to the super function we don’t need to pass it anything from the parent class. Don’t you agree this is much easier?

Now, keep in mind most classes will also have arguments passed to them. The super function will change even more when that happens.

It will look like the following:

     class MyParentClass():
          def __init__(self, x, y):
               pass

     class SubClass(MyParentClass):
          def __init__(self, x, y):
               super().__init__(x, y)
 
Again, this process is much more straightforward than the traditional method. In this case, we had to call the super function’s __init__ method to pass our arguments.

#### What Is the Super Function for Again?
The super function is extremely useful when you’re concerned about forward compatibility. By adding it to your code, you can ensure that your work will stay operational into the future with only a few changes across the board.

Ultimately, it eliminates the need to declare certain characteristics of a class, provided you use it correctly.

In order to use the function properly, the following conditions must be met:

The method being called upon by super() must exist
Both the caller and callee functions need to have a matching argument signature
Every occurrence of the method must include super() after you use it
You might start with a single inheritance class, but later, if you decide to add another base class – or more – the process goes a lot more smoothly. You only need to make a few changes as opposed to a lot.

There has been talk of using the super function for dependency injection, but we haven’t seen any solid examples of this – at least not practical ones. For now, we’re just going to stick with the description we’ve given.

Either way, now you’ll understand that super isn’t as bad as other devs purport it to be.
