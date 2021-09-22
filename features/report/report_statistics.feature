Ability: Report statistics
    At the end of a test run getting a summary about how many examples and steps
    were executed and how many of them passed and failed helps to get an
    overview if the test run was successful. To decide where to put some effort
    to make the next test run faster, it is also good to know how long the test
    took.


    Rule: statistics should appear as formatted text

      Example: verbose report
        Given the report format is "verbose"
          And a test run took 22 ms
          And one example in this run had 3 passed, 1 failed, 1 missing and 2 pending steps
         When the statistics are reported
         Then the resulting report output is
            """
              7 steps (3 passed, 1 failed, 1 undefined, 2 pending)
              0m 0.022s
            """

      Example: zero steps
