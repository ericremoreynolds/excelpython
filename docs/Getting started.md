# Getting started

## Description

xlpython is a lightweight, open source interface for interacting with Python from Excel. It enables easy access to Python scripts from Excel VBA, allowing you to substitute VBA with Python for complex automation tasks which would be facilitated by Python's extensive standard library. Features include:
* It is easily distributable, and does not require admin rights to install
* It is highly configurable, enabling you to specify exactly which installed Python distribution you want to target and how you want it to be run
* You can use any Python module from VBA without changing the original Python code
* Python runs out of process, which means that your Python code will behave exactly as it does when run from the command line
* The only dependency is PyWin32

## Installation

Unzip xlpython somewhere and import `xlpython.bas` into your workbook. Update the path to where you unzipped xlpython, and to test if it works try typing the following into the Immediate Window:
```
?Py.Str(Py.List(1, 2, 3))
```
Pressing enter, you should see the output `[1, 2, 3]`.

Now you can follow the tutorials showing how to manipulate Python objects VBA:

* [Usage 1](tutorials/Usage01.md)
* [Usage 2](tutorials/Usage02.md)
* [Usage 3](tutorials/Usage03.md)
* [Usage 4](tutorials/Usage04.md)

or you can delve deeper into how to target a particular Python installation and working environment:

* [Configuration](tutorials/Configuration01.md)

or you can learn how to analyse the problem when something goes wrong:

* [Troubleshooting](tutorials/Troubleshooting01.md)
