Ability: Report statistics
    At the end of a test run getting a summary about how many examples and steps
    were executed and how many of them passed and failed helps to get an
    overview if the test run was successful. To decide where to put some effort
    to make the next test run faster, it is also good to know how long the test
    took.


    Rule: statistics should appear when reported

      Example: verbose report
        Given the report format is set to verbose
          And statistics are 2 failed example each with a passed, failed and skipped step
          And statistics are 1 passed example with 3 passed steps
          And statitics contains 0.022s for the test duration
         When the statistics are reported
         Then line 1 of the resulting output is "3 scenarios (2 failed, 1 passed)"
          And line 2 of the resulting output is "9 steps (2 failed, 2 skipped, 5 passed)"
          And line 3 of the resulting output is "0m0.022s"
