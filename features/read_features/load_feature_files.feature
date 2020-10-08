Ability: load feature files
    Senfgurke will import features from feature files
    containing feature name and description, rules and examples as text.

Rule: search for feature files in the features dir beneath the current dir of the document containing the step implementation

    Example: derive feature from document dir
      Given the current application is stored under "/Users/cuke/source/senfgurke"
      When the location for features is requested
      Then the feature dir is set to "/Users/cuke/source/senfgurke/features"

Rule: the content of all files with the suffix .feature should be loaded as text, other files should be ignored

    Example: two feature files on top level
      Given two feature files in the feature dir
      When the feature dir is read
      Then two features are loaded

    Example: one feature file and one plain text file on top level
      Given one feature file and one text file in the feature dir
      When the feature dir is read
      Then one feature is loaded

    Example: one feature file on a sub-level
      Given one feature file located at a sub directory under the feature dir
      When the feature dir is read
      Then one feature is loaded

Rule: return error message if feature files are not accessible
