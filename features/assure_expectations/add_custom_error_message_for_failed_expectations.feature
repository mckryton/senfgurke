Ability: add custom error message for failed expectations
    Senfgurke offer different methods for comparing actual and expected
    values. When those expectations aren't met, both values are displayed in an
    error message. But sometimes just showing the expected and the actual value
    for a failed comparison is not expressive enough. To explain the intention
    in a more detailed way, all comparison methods will accept a custom error
    message as an additional optional parameter. When the expectation fails the
    custom error message will appended to the default error message to give
    further context.

  Rule: extend error message for failed expectations when custom error message is set

    Example: to_be expectations fail with custom error message
      Given an expected value was defined as <expected_value>
        And the actual value was evaluated as <actual_value>
        And a custom error message was defined as <custom_error_message>
       When expected and actual value are being compared using <comparison_type>
       Then <custom_error_message> is added to the expectation error message

      Examples:
        | comparison_type | expected_value | actual_value | custom_error_message                   |
        | to_be           |            200 |          400 | "cell doesn't have the expected width" |
        | not_to_be       |            400 |          400 | "width is still set to default size"   |

    @vba
    Example: has_member expectation fails with custom error message
      # VBA offers either a built-in Array data type or Collection objects to manage lists

      Given a <list_type> was used to save cell background color names "blue", "yellow" and "red"
        And a custom error message was defined as "no cell was marked as green to indicate success"
       When the expectation that the <list_type> contains "green" as member is validated
       Then "no cell was marked as green to indicate success" is added to the expectation error message

      Examples: list types provided by VBA
        | list_type  |
        | collection |
        | array      |

    Example: no custom error message is defined for a failed scenario



#TODO: allow whitespace formatting \n and \t in custom error messages
