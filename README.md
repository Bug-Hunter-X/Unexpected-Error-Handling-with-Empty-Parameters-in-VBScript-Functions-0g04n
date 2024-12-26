# Unexpected Error Handling with Empty Parameters in VBScript Functions

This repository demonstrates a common, yet subtle, error in VBScript related to handling empty or missing parameters passed to functions.  The `bug.vbs` file shows a function that attempts to check for empty parameters, but could still fail under certain conditions. The `bugSolution.vbs` file provides a more robust approach.

## Issue
VBScript's `IsEmpty` function, while useful, can sometimes behave unexpectedly when dealing with variants that might hold empty strings or null values.  The simple `If IsEmpty(...)` check isn't always sufficient for comprehensive parameter validation, leading to runtime errors that can be difficult to track down.

## Solution
The solution focuses on using more explicit checks to handle various forms of empty or null parameters.