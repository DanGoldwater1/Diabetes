# Example gptables pipeline
Simple worked example to demonstrate taking a csv file to an Excel file using the [gptables](https://github.com/best-practice-and-impact/gptables) package. This is a package designed to help us make xlsx files in accorance the ONS [best practice](https://gss.civilservice.gov.uk/about-us/support-for-the-gss/) guide. 

The result is stored in `Diabetes_data_gptables`. The target, existing output is the `National Audit` file. Note that although the data and content of the two are the same, the formatting does differ. Also note that the frontmatter on the cover page suffers from some typos and bugs - these are artifacts of what happens when we convert between different text formats, and serve as a reminder that text formatting should be checked on the output Excel file. 

The documentation for the `gptables` package is [here](https://gptables.readthedocs.io/en/latest/index.html). 