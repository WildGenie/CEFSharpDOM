# CEFSharpDOM

This is just an example of how to get a useable DOM from CEFSharp.

It's not fully working code, just a CEFSharp control with Document and Element classes.

The Document contains an List of Elements contained in the DOM, referenced by sourceIndex

A MutationObserver updates the List of elements, when the DOM changes.

Methods and Properties on the Elements call into the DOM when needed

To query the DOM, all threads need to be running, so you can't use hovertips on breakpoints.
