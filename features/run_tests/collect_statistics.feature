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


  Rule: steps and their enclosing feature sections should be counted

    Example: collect statistics for a feature with a single step example
      Given a feature
        """
          Feature: sample
            Example: single step scenario
              Given a missing step
        """
       When the feature is executed and statistics are collected
       Then the following statistics are collected
          | unit_type | unit_count |
          | features  | 1          |
          | rules     | 0          |
          | examples  | 1          |
          | steps     | 1          |


  #ToDo: add examples
  Rule: the result for examples and rule is passed or the result from the last failed step


  #ToDo: move this rule to the report feature
  Rule: Steps should be counted and grouped by their results

    Example: successful steps
      Given an example with 3 steps was executed successful
       When the results for the steps are summed up
       Then statistics results for steps is "3 steps (3 passed)"
