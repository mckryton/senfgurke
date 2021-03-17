# Features

* [Collect statistics](collect_statistics.feature)

  Several statistics are collected during a test run to learn more about
  the success of the test run in general. E.g. duration will tell you about
  the performance of the code under tests. Counting the excuted steps and
  their results will help you to get an overview about the test run results.

## Confirm Expectations

* [Confirm collection members](confirm_expectations/confirm_collection_members.feature)

  Senfgurke offers expectation functions to validate if a collection contains
  certain members or not.

* [Make steps executable](make_steps_executable.feature)

  Whenever a new step is added to a feature, Senfgurke will offer a matching
  step implementation function. Senfgurke tries to make the function name as
  similar to the step name as possible so that it's easy to identify the
  matching step implemenation for any step.
  But Function names in any programming language have to follow conventions.
  Conventions for function and variable names in VBA are documented here:

## Read Features

* [Load feature files](read_features/load_feature_files.feature)

  Senfgurke will import features from feature files
  containing feature name and description, rules and examples as text.

### Parse Features

* [Parse comments](read_features/parse_features/parse_comments.feature)

  To explain the intention of a feature it is possible to add a comment at any
  position. Any line starting with a # is considered to be such a comment. 

* [Parse examples](read_features/parse_features/parse_examples.feature)

  Examples (aka scenarios) are the executable part of a feature. An example

* [Parse features](read_features/parse_features/parse_features.feature)

  A Feature describes the functionality of a software that will solve some
  specific problem of it's user. It does so by giving examples of how the
  function works.
  The content of a featured follows a nested structure. A feature can contain
  one or more rules (see parse rules for more). While a rule can contain one
  or more examples (see parse examples for more).
  Simple features can contain just some few examples without any rule.

* [Parse rules](read_features/parse_features/parse_rules.feature)

  A rule describes the logic or policy a feature shold follow. All examples
  following a rule should explain this policy. A rule is limited by the start
  of the next rule or by the end of the feature.

* [Parse steps](read_features/parse_features/parse_steps.feature)

  Steps are the building blocks of examples. Every step has a matching step
  definition where the step is expressed as working code. As result an example
  can be considered as working if all the code from the step definition is
  executed without errors.

#### Parse Steps

* [Parse docstrings](read_features/parse_features/parse_steps/parse_docstrings.feature)

  Sometimes steps refer to a larger text strucure having multiple lines (e.g.
  describing a feature in a single step). Therefor step can be expanded by one
  docstring following the step directly. Docstrings are embraced by three double
  quotation marks

* [Parse step expressions](read_features/parse_features/parse_steps/parse_step_expressions.feature)

  To match a step with it's step definition it has to be disassembled into its

* [Parse tags](read_features/parse_features/parse_tags.feature)

  Tags will help you to navigate through features and examples. For example it
  is possible to run only examples marked with a specific tag.
  Examples will inherit their tags from the feature. If a tag is set for a
  feature, all examples will have this tag too.

## Report

* [Report in progress format](report/report_in_progress_format.feature)

  When running examples for all features, a verbose output would be confusing.
  Therefore the progress format will mark successul executed examples with
  a single dot to indicate the execution progress. Different results can be
  detected by a matching letter (e.g. F for failed steps).

  Background:
  Given the report format is "progress"

* [Report in verbose format](report/report_in_verbose_format.feature)

  While developing new features or debugging selected examples the verbose
  report format come in handy. It will show step definition next to the latest
  execution result as well as name and descriptions for features and rules.

  Background:
  Given the report format is "verbose"

* [Report statistics](report/report_statistics.feature)

  At the end of a test run getting a summary about how many examples and steps
  were executed and how many of them passed and failed helps to get an
  overview if the test run was successful. To decide where to put some effort
  to make the next test run faster, it is also good to know how long the test
  took.

## Run Features

* [Run examples](run_features/run_examples.feature)

  Examples are the building blocks for describing the application under test.
  Running examples means that Senfgurke will try to find and execute a
  matching function for each step in an example (aka step implementation).
  The result of the executed step will tell you if the application works as
  expected (see feature "run steps" for more).

* [Run features](run_features/run_features.feature)

  A features represents the business intention on a more abstract level. The
  feature description should tell you why the feature was implemented in the
  first place.
  Senfgurke runs feature by finding all the examples (see run examples
  feature for more) and executing all the steps from those examples (see run
  steps feature for more).

* [Run steps](run_features/run_steps.feature)

  Steps are the glue that connects features with code. Every step can have a
  matching function (step implementation). Executing those functions will

## Support Functions

* [Get unix time](support_functions/get_unix_time.feature)

  There is no built in support for unix time stamps in vba. However using unix
  time stamps helps to calculate durations.
