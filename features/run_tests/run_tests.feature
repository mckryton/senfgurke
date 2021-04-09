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


  Rule: for any given tag execute only those features assigned with the tag

    Example: matching tag assigned to a feature
      Given a feature tagged with "@sample_tag" with one example
        And an un-tagged feature with one example
       When a test is started with "@sample_tag" as parameter
       Then only the tagged feature is executed

    Example: non-matching tag assigned to a feature


    Example: no tags assigned to a feature
