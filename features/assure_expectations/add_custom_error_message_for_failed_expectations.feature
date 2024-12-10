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

    @wip
    Example: text comparison expectations fail with custom error message
      Given an text was expected to include <expected_text>
        But the actual text was <actual_text>
        And a custom error message was defined as <custom_error_message>
       When the text is compared using <comparison_type>
       Then <custom_error_message> is added to the expectation error message

      Examples:
        | comparison_type | expected_text  | actual_text                          | custom_error_message                                       |
        | starts_with     |  "Error:"      | "Warning: address is out of bounds"  | "This kind of scenario must result in an error"            |
        | ends_with       |  "200"         | "Web service returned status 404"    | "Calling the web service should return status 200"         |
        | includes_text   |  "status"      | "Web service is not responding"      | "Calling the web service should return the current status" |

    @vba-specific
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
