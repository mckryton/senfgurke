Ability: run tests
  Executing or running tests using Senfgurke happens in 2 steps. First Senfgurke
  will read all the features form text files using the suffix .feature. It will
  turn the structured content of those files into an executable setup by
  interpreting (parsing) the Gherkin language of the feature files.

  Rule: any syntax error found when parsing feature files should be reported
    See the report_in_<format_type>_format features under /features/report
    to find out how the error messages is displayed.

    Example: Feature doesn't start with a feature keyword
      Given a feature was loaded from a file "example.feature"
        """
          this is some random text
          and not a feature
        """
       When Senfgurke executes the feature
       Then the error "Feature lacks feature keyword at the beginning" is reported
        And the name of the feature file is reported as location of the error


  Rule: tests started with tag as parameter should run only examples with matching tags

    Example: example inherits tag from feature
      Given a feature was loaded as
        """
          @sample_tag
          Feature: tagged feature
            Example: sample from tagged feature
              Given a step
        """
        And a feature was loaded as
        """
          Feature: un-tagged feature
            Example: sample from un-tagged feature
              Given a step
        """
       When a test is started with "@sample_tag" as parameter
       Then only the example from the tagged feature was executed

     Example: example inherits tag from rule
       Given a feature was loaded as
         """
           Feature: tagged feature
             Example: example assigned to the feature
               Given a step

             @sample_tag
             Rule: tagged rule
               Example: example assigned to the rule
                 Given a step

             Rule: untagged rule
               Example: some other example
                 Given a step
         """
        When a test is started with "@sample_tag" as parameter
        Then only the example from the tagged rule was executed

    Example: tag doesn't match
      Given a feature was loaded as
        """
          @one
          Feature: tagged feature
            @two
            Example: sample from tagged feature
              Given a step
        """
       When a test is started with "@three" as parameter
       Then no example was executed

#TODO: explain all available parameters for running tests
