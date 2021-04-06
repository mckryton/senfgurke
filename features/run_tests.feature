Ability: run tests
  Executing or running tests using Senfgurke happens in 2 steps. First Senfgurke
  will read all the features form text files using the suffix .feature. It will
  turn the structured content of those files into an executable setup by
  interpreting (parsing) the Gherkin language of the feature files.


  Rule: any syntax error found when parsing feature files should be reported
    See the report_in_<format_type>_format features under /features/report
    to find out how the error messages is displayed.

    Example: Feature doesn't start with a feature keyword
      Given a feature starting with random text in "sample.feature"
       When Senfgurke executes the feature
       Then the error "Feature lacks feature keyword at the beginning" is reported
        And the name of the feature file is reported as location of the error
