Ability: cleanup after example run
    Running examples will sometimes leave some artefacts created specifically
    for the example. For example one step will require a new document to
    demonstrate some functionality on this document. This can cause a lot of new
    documents of the end of the test. Therefore Senfgurke provides a so called
    hook function to allow you to add some cleanup code (e.g. for closing the
    documents created by any example). Senfgurke will call this function every
    time an example is finished.

    Access to any hook function requires to steps:
      1. add a new variable to the step implementation class:
            Private WithEvents ExecutionHooks As TExecutionHooks
      2. assign this variable to your test sessions
            set ExecutionHooks = TRun.Session.ExecutionHooks
    Execution hook function will now appear in the VB editor in the class
    window when the ExecutionHooks entry is selected

    Background:
       Given a scenario
         """
           Given a car is filled with gas
           When the engine is started
           Then the car is ready to drive
         """

    Rule: call the after scenario function whenever a scenario was finished
      The scenario status should be attached to the event to allow further
      analysis in case of failure.
      Note: scenario and example are synonyms

      Example: raise an "after scenario" event for an example finished successfully
        When the scenario succeeds successfully for all steps
          """
            OK      Given a car is filled with gas
            OK       When the engine is started
            OK       Then the car is ready to drive
          """
        Then the after scenario hook function is called after the example
         And the execution status of the scenario is attached as "OK"

      Example: raise an "after scenario" event for an example with incomplete execution
        When the scenario fails because of missing step implementations
          """
            MISSING  Given a car is filled with gas
            SKIPPED   When the engine is started
            SKIPPED   Then the car is ready to drive
          """
        Then the after scenario hook function is called after the example
         And the execution status of the scenario is attached as "FAIL"

#TODO: add examples for outline scenarios
