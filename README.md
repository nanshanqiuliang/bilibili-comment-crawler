# bilibili_comments- crawler
A crawler for bilibili. 
It asks you to provide your cookies. You'd better not use your main account to avoid probabe loss.
Just follow the instructions.
Only the main comments will be crawled, which means the comments in response to a comment will not be recorded. This function is under developing.
The .xlsx file is an example

2026/5/28
Added a function to cache the procedure. Avoid the loss of all the data when encountering accident. Enable users to pause the program and later on to continue from the pause.

2026/5/29
Crawling replies under main comments is enabled. Users can set a threshold to decide which comment deserves reading the reply. i.e. For those comments that have less than 5 replies, it is a waste of time to read them since they contribute little to the total number of comments.
