# Stock-analysis-practice

Hello guys
{
	"version": "0.1.0",

	// The command is tsc. Assumes that tsc has been installed using npm install -g typescript
	"command": "python",

	// The command is a shell script
	"isShellCommand": true,

	// Show the output window only if unrecognized errors occur.
	"showOutput": "always",

	// args is the HelloWorld program to compile.
	"args": ["${file}"],

	// use the standard tsc problem matcher to find compile problems
	// in the output.
	"problemMatcher": "$tsc"
}