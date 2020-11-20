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

  Senfgurke will read feature specs from text
  and identify examples and their example steps from within features

* [Parse features](read_features/parse_features/parse_features.feature)

  Senfgurke will read feature specs from text
  and identify it's elements like descriptions, rules and examples

* [Parse rules](read_features/parse_features/parse_rules.feature)

  Senfgurke will read feature specs from text
  and identify rules as counterpart for the following examples

* [Parse tags](read_features/parse_features/parse_tags.feature)

  Senfgurke will read feature specs from text
  and identify tag assigned to the whole features or just examples

## Report

* [Report in verbose format](report/report_in_verbose_format.feature)

  While executing examples senfgurke will send messages about progress
  and success for reporting. The verbose formatter will turn those messages into
  a verbose report that is printed on the debug console.

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
