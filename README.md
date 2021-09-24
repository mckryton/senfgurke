# Senfgurke
Senfgurke is example driven test framework for VBA. What does this mean? Using Senfgurke you can turn examples given in natural language into automated tests run by VBA ([Visual Basic for Applications](https://docs.microsoft.com/en-us/office/vba/api/overview/)).

**BEWARE!** This is work in progress. Future versions might break your test automation code from older versions!

For example someone asks you to write a new sum function wich adds the value of 1 to the result by giving you this example:

```
Example: add +1 to sum
  Given a is 2
    And b is 3
   When sum+1 is applied to a and b
   Then the result is 5
```

You add this example to a feature and save everything in file named 'sum_plus_one.feature' to a directory named 'features'. The directory should be in the same location as your office file containing your VBA code.

If you run Senfgurke the first time it will suggest you to add a new function like this:

```
Public Sub Given_a_is_INT_C722764574FB(step_parameters As Collection)
    'Given a is 2

End Sub
```

Please don't get confused with the C722764574FB part of the function name. This hash value helps Senfgurke to match the function with the original step from the example. Now it's up to you to fill the function with code, so that you can test your new function to be sure that sum+1 returns the correct results.  

When you have repeated this for every step of the example and run Senfgurke you might receive this on the console:

```
Feature: sum plus one


  Rule: add one to sum results

    Example: add +1 to sum
     OK       Given a is 2
     OK       And b is 3
     OK       When sum+1 is applied to a and b
     OK       Then the result is 5
```

This way Senfgurke tells you if your code was successful or caused any error.

## Setup (short version)
1. create a new office file and add some VBA
2. add all files from the /source folder from Senfgurke's repository.
3. create a 'features' directory at the same location from your office file
4. create a (Gherkin)[https://cucumber.io/docs/gherkin/] feature file (only basic elements will work)
5. go to your code and add a new class 'step_definition_<your_feature>'
6. go to the StepImplementations Property in the TConfig module and add your class
7. re-comile your project
8. call Trun.test
