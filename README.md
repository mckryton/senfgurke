# Senfgurke
Senfgurke is example driven test framework for VBA. What does this mean? Using Senfgurke you can turn examples given in natural language into automated tests run by VBA ([Visual Basic for Applications](https://docs.microsoft.com/en-us/office/vba/api/overview/)).

- [Introduction](#Introduction)
- [Setup](#Setup)
- [Functional design](#Functional-design)


**BEWARE!** This is work in progress. Future versions might break your test automation code from older versions!

## Introduction
Imagine someone is asking you to write a new "special" sum function for Excel wich adds the value of 1 to the result by giving you this example:

```
Example: add +1 to sum
  Given a is 2
    And b is 3
   When sum+1 is applied to a and b
   Then the result is 6
```

You may now add this example to a feature and save everything in file named 'sum_plus_one.feature' to a directory named 'features'. The directory should be in the same location as your office file containing your VBA code.

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

## Setup
Senfgurke is provided as an application specific Addin. This way you can easily separate the application under test and the Senfgurke test framework. To start a new VBA application follow these steps:

1. Activate the Senfgurke Add-In for your application
2. Create a new document for your code (eg.: a new xlm file for Excel)
3. Import the TRun module from step_definitions/TRun.bas to your application
4. Add a reference to the Senfgurke Add-In for your application in the VBA IDE
5. Create a new directory named "features" at the same directory of your application (the document containing your VBA code)
6. Create your first Gherkin feature file under features (add at least 1 example)
7. Execute TRun.test from the Immediate Window in your VBA IDE
8. Add a new Class "Step_<your_feature>" to your  application
9. Copy the suggested code for the step functions from the Immediate Window to your Step class
10. Register your Step class by extending the StepImplementations Property in the TRun module

## Functional design
The following [event map](https://vimeo.com/130202708) will explain what will happen when you ask Senfgurke to execute your features.
![event map for Senfgurke](https://raw.githubusercontent.com/mckryton/senfgurke/master/design/senfgurke%20key%20events.svg "Senfgurke key events")

### [Run a test](features/run_tests/run_tests.feature)
Tests are usually started from the VBA console window. This way you can add tags or filters (for file names) to restrict the test run to specific tags or feature files.

### [Load features files](features/read_features/load_feature_files.feature)
The first thing Senfgurke will do is looking for feature files and loading them into memory for later processing.

### [Parse features](features/read_features/parse_features.feature)
Having all the features in memory makes it easier for Senfgurke to translate (parse) the [Gherkin language](https://cucumber.io/docs/gherkin/) into detailed instructions for later execution. E.g., examples without matching tags set on the test start might be ignored for later execution or [background steps](https://cucumber.io/docs/gherkin/reference/#background) from a feature will be added to every example (aka scenario) in this feature.

### [Run features](features/run_tests/run_features.feature)
The next step is obviously about executing all those detailed execution instructions from the step before. This will also include returning all the results from the execution.

### [Report results](features/report/report_in_verbose_format.feature)
In parallel to the execution of the features mentioned above, results will be reported in different formats. Default is verbose format which writes the Gherkin features to the VBA console window by just adding the execution result to each step of the examples.

### [Report statistics](features/report/report_statistics.feature)
At the end of every test run Senfgurke will add some statistics, for example duration and number of executed example steps.
