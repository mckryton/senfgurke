# Features

## Assure Expectations

* [Check if objects exists](assure_expectations/check_if_objects_exists.feature)

  VBA distinguishes between variables assigned to objects and variables
  assigned to basic data types (e.g. string, integer, boolean). This
  expectation validates if a variable is assigned to an object or not. The
  expectation is using the Nothing keyword to determin if a variable is
  assigned to an objoct or not.

* [Confirm collection members](assure_expectations/confirm_collection_members.feature)

  A collection in VBA is like a named array (see

## Read Features

* [Load feature files](read_features/load_feature_files.feature)

  One of the essential goals of BDD is to support collaboration between
  stakeholders and developer. So to keep barriers as low as possible business
  input is accepted as plain text. Therefore Senfgurke will import features
  from plain text files using the .feature suffix.

* [Parse features](read_features/parse_features.feature)

  A feature describes the functionality of a software that will solve some
  specific problem of it's user. It does so by giving examples of how the
  function works.
  For a better understanding a featured follows a nested structure. A feature
  can contain zero or more rules (see parse rules for more). While a rule can
  contain zero or more examples (see parse examples for more). Of course only
  examples can be executed. So a feature or or rule without any assigned
  example will have no further effect on execution.
  Simple features can contain just some few examples without any rule.

### Parse Features

* [Parse comments](read_features/parse_features/parse_comments.feature)

  To explain the intention of a feature it is possible to add a comment at any
  position. Any line starting with a # is considered to be such a comment.

* [Parse examples](read_features/parse_features/parse_examples.feature)

  Examples (aka scenarios) are the executable part of a feature. An example

* [Parse feature definition](read_features/parse_features/parse_feature_definition.feature)

  To distinct feature file from other text files a feature files shows its name
  on the top. Only tags, comments and empty lines are allowd above the feature
  name. Any description has to put below the feature name.
  The feature name is marked with any of the keywords reserved by Gherkin for
  features like feature, Abilty or Business Needs and a colon.

* [Parse rules](read_features/parse_features/parse_rules.feature)

  Quite often a feature has more than one aspect. E.g. most feature will follow
  several business policies. A rule will help you to reflect those aspects in
  the feature description. A rule describes the logic or policy a feature should
  follow. All examples following a rule should explain this policy.

* [Parse steps](read_features/parse_features/parse_steps.feature)

  Steps are the building blocks of examples. Every step has a matching step
  definition where the step is expressed as working code. As result, an example
  can be considered "working" if all the code from the step definition is
  executed without errors.

#### Parse Steps

* [Parse docstrings](read_features/parse_features/parse_steps/parse_docstrings.feature)

  Sometimes steps refer to a larger text strucure having multiple lines (e.g.
  describing a feature in a single step). Therefore steps can be expanded by one
  docstring following the step directly. Docstrings are embraced by three double
  quotation marks

* [Parse outline parameters](read_features/parse_features/parse_steps/parse_outline_parameters.feature)

  Examples can include an outline table providing one or more sets of
  parameters to vary the excution of the example. For this step are allowed to
  use placeholder to fill in outline parameters during execution.

* [Parse step expressions](read_features/parse_features/parse_steps/parse_step_expressions.feature)

  To match a step with it's step definition it has to be disassembled into its

* [Parse tables](read_features/parse_features/parse_tables.feature)

  Tables are a good way to keep examples concise and compact. Tables can be used
  as data table to extend steps or as outlines to extend examples.

#### Parse Tables

* [Parse data table](read_features/parse_features/parse_tables/parse_data_table.feature)

  If examples require more extensive pre-conditions, brevity of the Given steps
  can be improved by using a data table instead of a lot of single steps.

* [Parse outlines](read_features/parse_features/parse_tables/parse_outlines.feature)

  Sometimes examples differ only in a small set of parameters. So instead of
  duplicating those examples for each set of parameter it is possible to add a
  table providing a set of parameters with each row at the end of a single
  example.

* [Parse tags](read_features/parse_features/parse_tags.feature)

  Tags help you navigate through features and examples. For example it is
  possible to run only examples marked with a specific tag.
  Examples inherit all tags from their features. If a tag is set for a
  feature, all examples in that feature will have this tag too.
  Any line starting with an @ sign is a tag line where tags are starting
  with @ and ending with space or linebreak.

## Report

* [Report in progress format](report/report_in_progress_format.feature)

  Running examples for all features would be confusing because of the amount
  of details for all the steps.
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

## Run Tests

* [Collect statistics](run_tests/collect_statistics.feature)

  Several statistics are collected during a test run to learn more about
  the success of the test run in general. E.g. duration will tell you about
  the performance of the code under tests. Counting the excuted steps and
  their results will help you to get an overview about the test run results.

  Background:
  Given a new test run started collecting statistics

* [Connect steps with step functions](run_tests/connect_steps_with_step_functions.feature)

  Whenever a new step is added to a feature, Senfgurke will offer a matching
  step function. Senfgurke tries to make the function name as similar to the
  step name as possible so that it's easy to identify the matching step
  implementation for any step. But function names in any programming language
  have to follow conventions. Conventions for function and variable names in
  VBA are documented here:

* [Run examples](run_tests/run_examples.feature)

  Examples are the building blocks for describing the application under test.
  Running examples means that Senfgurke will try to find and execute a
  matching function for each step in an example (aka step implementation).
  The result of the executed step will tell you if the application works as
  expected (see feature "run steps" for more).

### Run Examples

* [Run outline example](run_tests/run_examples/run_outline_example.feature)

  Examples can include outline tables to allow the repetition of an example but
  with a different set of parameters. So the example will be executed once for
  every row of the outline table while the value from each column will replace
  the corrosponding placeholder in the examples steps.

* [Run features](run_tests/run_features.feature)

  A features represents the business intention on a more abstract level. The
  feature description should tell you why the feature was implemented in the
  first place.
  Senfgurke runs features by finding all the examples (see the run examples
  feature for more) and executes all the steps from those examples (see the
  run steps feature for more).

* [Run rules](run_tests/run_rules.feature)

  _no description_

### Run Steps

* [Run steps](run_tests/run_steps/run_steps.feature)

  Steps are the glue that connects features with code. Every step can have a
  matching function (step implementation). Executing those functions will

* [Save variables in context](run_tests/run_steps/save_variables_in_context.feature)

  Typically when implementing step definition functions there is a need to
  access the values defined in Given steps in the step definition functions
  of the When and Then steps. But steps for a single example can have theire
  step definitions in different classes, for example if the same step appears
  in more than one example. So to share values between step definition
  functions Senfgurke provides a context object as parameter for each step
  definition function that keeps this variable during the execution of a
  single example.

* [Show step function templates](run_tests/run_steps/show_step_function_templates.feature)

  Whenever a step is executed and no matching step function is found Senfgurke
  will generate a code snippet for a new step function. This way it should be
  easy to assign code to the given step.

* [Run tests](run_tests/run_tests.feature)

  Executing or running tests using Senfgurke happens in 2 steps. First Senfgurke
  will read all the features form text files using the suffix .feature. It will
  turn the structured content of those files into an executable setup by
  interpreting (parsing) the Gherkin language of the feature files.

## Support Functions

* [Get unix time](support_functions/get_unix_time.feature)

  There is no built in support for unix time stamps in vba. However using unix
  time stamps helps to calculate durations.
