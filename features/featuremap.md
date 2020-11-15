# Senfgurke Features

## Features


### General

#### [Make steps executable](features/make_steps_executable.feature)

  Senfgurke will derive step function names from example steps so that a
  developer can add test code for each step.


### Read_features

#### [Load feature files](features/read_features/load_feature_files.feature)

  Senfgurke will import features from feature files
  containing feature name and description, rules and examples as text.


### Parse_features

#### [Parse comments](features/parse_features/parse_comments.feature)

  Senfgurke will read feature specs from text
  and identify comments

#### [Parse examples](features/parse_features/parse_examples.feature)

  Senfgurke will read feature specs from text and identify examples and their
  example steps from within features

#### [Parse tags](features/parse_features/parse_tags.feature)

  Senfgurke will read feature specs from text
  and identify tag assigned to the whole features or just examples

#### [Parse rules](features/parse_features/parse_rules.feature)

  Senfgurke will read feature specs from text
  and identify rules as counterpart for the following examples

#### [Parse features](features/parse_features/parse_features.feature)

  Senfgurke will read feature specs from text
  and identify it's elements like descriptions, rules and examples


### Confirm_expectations

#### [Confirm collection members](features/confirm_expectations/confirm_collection_members.feature)

  Senfgurke offers expectation functions to validate if a collection contains
  certain members or not.


### Run_features

#### [Run features](features/run_features/run_features.feature)

  Senfgurke will locate the matching step implementation for each step
  in a feature and execute it.

#### [Run steps](features/run_features/run_steps.feature)

  Senfgurke will locate the matching step implementation for each step
  of an example and execute it.

#### [Run examples](features/run_features/run_examples.feature)

  Senfgurke will locate the matching step implementation for each step
  of an example and execute it.

### Report


#### [Report in verbose format](features/report/report_in_verbose_format.feature)

  While executing examples Senfgurke will send messages about progress
  and success for reporting. The verbose formatter will turn those messages
  into a verbose report that is printed on the debug console.
