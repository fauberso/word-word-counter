# MS Word Word Counter

This is a small example Java program that counts the words in an MS Word document. It shows how to implement some complex requirements for word counting. 

The background is: For some academic papers I needed to write, the word count had to be calculated excluding:
* Cover page and TOC (i.e. anything before the "Introduction" heading)
* Appendices and Bibliography (i.e. anything after the "Bibliography" heading)
* Caption text for images (i.e. anything with a style named "Caption Text")
* Text boxes in the sidebar (counted and output separately in the example)
* Citation sources (implemented as: Anything in parentheses containing a 4-digit number)

The example program will output the word count, and the number of ignored words. The sum should be close to what Word reports.