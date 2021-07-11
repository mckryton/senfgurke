Ability: load feature files
    One of the essential goals of BDD is to support collaboration between
    stakeholders and developer. So to keep barriers as low as possible business
    input is accepted as plain text. Therefore Senfgurke will import features
    from plain text files using the .feature suffix.


  Rule: the feature directory should be a subdir of the application under test
    Search for feature files in the features dir beneath the current dir of
    the document containing the step implementation.

    Example: derive feature from document directory
      Given the current application is stored under "/Users/cuke/source/senfgurke"
       When the location for features is requested
       Then the feature dir is set to "/Users/cuke/source/senfgurke/features"


  Rule: only files using the suffix ".feature" should be loaded from the feature directory

    Example: two feature files on top level
      Given two feature files in the feature dir
       When the feature dir is read
       Then 2 feature(s) are loaded

    Example: one feature file and one plain text file on top level
      Given one feature file and one text file in the feature dir
       When the feature dir is read
       Then 1 feature(s) are loaded

    Example: one feature file on a sub-level
      Given one feature file located at a sub directory under the feature dir
       When the feature dir is read
       Then 1 feature(s) are loaded


  Rule: for a given filter only feature files with a matching name should be loaded
    Comparison starts from the left. Filter matches when the leftmost characters
    from the feature file name matches the filter.

    Example: file name matches filter completely
      Given a feature file "play.feature"
        And a feature file "plug.feature"
       When a test is started with "play" as filter for feature names
       Then only "play.feature" is loaded and executed

    Example: file names match filter partially
      Given a feature file "play.feature"
        And a feature file "plug.feature"
       When a test is started with "pl" as filter for feature names
       Then both features are loaded and executed


  Rule: return error message if feature files are not accessible
#    Example: feature directory is unavailable
#      Given feature dir is set to "/this/path/does/not/exist/features"
#      When senfgurke reads features from feature dir
#      Then the import returns an error message "can't access feature dir ><"
