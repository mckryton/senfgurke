@vba-specific @unit-test
Feature: get unix time
  There is no built in support for unix time stamps in vba. However using unix
  time stamps helps to calculate durations.


  Rule: a unix time stamps should be the number of milliseconds since 1.1.1970 00:00:00

    Example: date time conversion
      # vba is using the local timezone
      Given the current date is 01.01.1970 00:00:01
       When the date is converted to an unix timestamp
       Then the value of the unix timestamp is 1000
