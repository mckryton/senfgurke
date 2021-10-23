| date | decision |
|------|----------|
| 23.10.2021 | Step implementations and the Senfgurke framework should be kept in separate files. This way it's possible to include Senfgurke into a new project without including Senfgurke's own step implementation classes. |
| 11.11.2020 | If a step contains expressions (string, int or float parameter values), those parameters will be given to the step function as a collection object. This is because the CallByName function can't translate an array into single parameters. |
| 07.07.2020 | Senfgurke uses the debug console for output. This way it's not limited to a specific office application.|
| 07.07.2020 | ~~Senfgurke's functionality is limited. For simplicity it won't parse examples from external textfiles and will use text variables from feature classes.~~ fixed :-)|
| 07.07.2020 | The application should run on mac os as well as on windows -> it must not contain references to Windows only libraries (e.g. vbscript). Unfortunately this means it can't use regular expressions. :-/ |
| 26.09.2020 | add hashes to step implementation function to make them unique because VBA don't know annotations like in Java|
