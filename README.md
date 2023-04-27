# VBA-Challenge

None of the code was copied from any other sources. I did use ChatGPT to answer quick questions about why things weren't working here or there if I got an error and couldn't see why after attempts on my own. 

Once was to figure out how to establish the firstOpen variable, although I must have asked it incorrectly because it told me to do (i - 251) and I realized I needed to do (i - 250) because it was storing cell C1 and giving me a type mismatch. 
  
The other time ChatGPT directly assisted was in telling me that I was trying to run WorksheetFunction.Max("k:k") and that was a string so I had to fix it to .Max(ws.Range("k:k")) instead which was just a silly error.
  
The rest of the code was based off in-class activies or questions I asked of the instructor.
