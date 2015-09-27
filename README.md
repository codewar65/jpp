# jpp
JavaScript Preprocessor

Preprocessor for JavaScript to allow for C/C++ style preprocessor directives. Console application written in Visual Basic.

Usage:
- `jpp.exe <inputfile> -o <outputfile>`
- `jpp.exe < inputfile > outputfile`

Supported directives:
- `#pragma`
- `#include`
- `#define`
- `#undef`
- `#if`
- `#ifdef`
- `#ifndef`
- `#endif`
- `#error`
- `#//`

Built in define macros:
- `__LINE__`
- `__FILE__`
- `__DATE__`
- `__TIME__`

Pragma switches:
- `tab2`, `tab3`, `tab4`, `tab5`, `tab6`, `tab7`, 'tab8` - set tabstop for output source. defaults to tab8
- `addfrom` - append "// filename:linenumber" to end of source lines
- `stripemptylines` - remove blank lines
- `stripcomments` - remove comments
- `pack` - remove excessive whitespace from lines of source
- `minify` - run AjaxMini on source. will override all above switches
