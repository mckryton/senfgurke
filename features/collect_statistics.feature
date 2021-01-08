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


  Rule: Duration should be reported as "#m #.###s"

         Example: duration less than a second
           Given duration of a test run is 42 ms
            When the duration is calculated
            Then the resulting output is "0m 0.042s"

         Example: duration more than a minute
           Given duration of a test run is 74531 ms
            When the duration is calculated
            Then the resulting output is "1m 14.531s"
