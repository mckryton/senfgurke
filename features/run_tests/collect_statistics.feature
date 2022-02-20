Ability: Collect statistics
    Several statistics are collected during a test run to learn more about
    the success of the test run in general. E.g. duration will tell you about
    the performance of the code under tests. Counting the excuted steps and
    their results will help you to get an overview about the test run results.

  Background:
    Given a new test run started collecting statistics

  Rule: collect start and end time for a test run

    Example: calculate test run duration
      Given a feature
        """
          Feature: sample
            Example: single step scenario
              Given a missing step
        """
       When the feature is executed and statistics are collected
       Then start and end time for the test run are set


  Rule: duration should be calculated as text "#m #.###s"

     Example: duration less than a second
       Given duration of a test run is 42 ms
        When the duration is calculated
        Then the resulting output is "0m 0.042s"

     Example: duration more than a minute
       Given duration of a test run is 74531 ms
        When the duration is calculated
        Then the resulting output is "1m 14.531s"


  Rule: all types shold be summed up while steps should also be counted by status

    Example: successful steps
      Given a feature
         """
           Feature: sample
             Example: single step scenario
               Given a valid step
                 And a valid step
                 And a valid step
         """
       When the feature is executed and statistics are collected
        And the results for the steps are summed up
       Then statistics results for steps is "3 steps (3 passed)"

    Example: collect statistics for a feature with a single step example
      Given a feature
         """
           Feature: sample
             Example: single step scenario
               Given a undefined step
         """
       When the feature is executed and statistics are collected
        And the results for the steps are summed up
       Then statistics results for steps is "1 step (1 undefined)"
