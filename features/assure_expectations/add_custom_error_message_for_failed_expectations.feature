Ability: add custom error message for failed expectations
    Senfgurke offer different methods for comparing actual and expected
    values. When those expectations aren't met, both values are displayed in an
    error message. But sometimes just showing the expected and the actual value
    for a failed comparison is not expressive enough. To explain the intention
    in a more detailed way, all comparison methods will accept a custom error
    message as an additional optional parameter.

  Rule: extend error message for failed expectations when custom error message is available
@wip
    Example: comparison fails with custom error message
      Given an expected value is defined as <expected_value>
        And the actual value is <actual_value>
        And the custom error message is defined as <custom_error_message>
       When expected and actual value are being compared using <comparison_type>
       Then <custom_error_message> is added to the expectation error message

      Examples:
        | comparison_type | expected_value | actual_value | custom_error_message                   |
        | to_be           |            200 |          400 | "cell doesn't have the expected width" |
        | not_to_be       |            400 |          400 | "width is still set to default"        |

    Example: no custom error message is defined for a failed scenario



#TODO: allow whitespace formatting \n and \t in custom error messages
