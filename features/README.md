# Features

## Confirm Expectations

* [Confirm collection members](confirm_expectations/confirm_collection_members.feature)

  Senfgurke offers expectation functions to validate if a collection contains
  certain members or not.

* [Make steps executable](make_steps_executable.feature)

  Senfgurke will derive step function names from example steps so that a
  developer can add test code for each step.

## Read Features

* [Load feature files](read_features/load_feature_files.feature)

  Senfgurke will import features from feature files
  containing feature name and description, rules and examples as text.

### Parse Features

* [Parse comments](read_features/parse_features/parse_comments.feature)

  Senfgurke will read feature specs from text
  and identify comments

* [Parse examples](read_features/parse_features/parse_examples.feature)

  Examples (aka scenarios) are the executable part of a feature. An example

* [Parse features](read_features/parse_features/parse_features.feature)

  Senfgurke will read feature specs from text
  and identify it's elements like descriptions, rules and examples

* [Parse rules](read_features/parse_features/parse_rules.feature)

  Senfgurke will read feature specs from text
  and identify rules as counterpart for the following examples

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

  Senfgurke will read feature specs from text
  and identify tag assigned to the whole features or just examples

## Report

* [Report in progress format](report/report_in_progress_format.feature)

  When running examples form all features a verbose output would be confusing.
  Therefore the progress format will mark successul executed examples with
  a single dot to indicate the execution progress. Different results can be
  detected by a matching letter (e.g. F for failed steps).

  Background:
  Given the report format is set to progress

* [Report in verbose format](report/report_in_verbose_format.feature)

  While developing new features or debugging selected examples the verbose
  report format come in handy. It will show step definition next to the latest
  execution result as well as name and descriptions for features and rules.

  Background:
  Given the report format is set to verbose

## Run Features

* [Run examples](run_features/run_examples.feature)

  Senfgurke will locate the matching step implementation for each step
  of an example and execute it.

* [Run features](run_features/run_features.feature)

  Senfgurke will locate the matching step implementation for each step
  in a feature and execute it.

* [Run steps](run_features/run_steps.feature)

  Senfgurke will locate the matching step implementation for each step
  of an example and execute it.
