Ability: Collect statistics
    Several statistics are collected during a test run to learn more about
    the success of the test run in general. E.g. duration will tell you about
    the performance of the code under tests. Counting the excuted steps and
    their results will help you to get an overview about the test run results.


  Rule: collect start and end time for a test run

    Example: calculate test run duration
      Given a feature description with a step
       When the feature is executed and statistics are collected
       Then start and end time for the test run are set


  Rule: Duration should be calculated as text "#m #.###s"

     Example: duration less than a second
       Given duration of a test run is 42 ms
        When the duration is calculated
        Then the resulting output is "0m 0.042s"

     Example: duration more than a minute
       Given duration of a test run is 74531 ms
        When the duration is calculated
        Then the resulting output is "1m 14.531s"


  Rule: Steps should be counted with their results

    Example: single step
      Given a feature with an example with a missing step
       When a test runs only this feature
       Then one step with it's result was counted


  Rule: Steps should be counted and grouped by their results

    Example: successful steps
      Given an example with 3 steps was executed successful
       When the results for the steps are summed up
       Then statistics results for steps is "3 steps (3 passed)"
