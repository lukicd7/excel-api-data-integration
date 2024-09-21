#
# This is a Plumber API. You can run the API by clicking
# the 'Run API' button above.
#
# Find out more about building APIs with Plumber here:
#
#    https://www.rplumber.io/
#

library(plumber)

#* @apiTitle Plumber Example API
#* @apiDescription Plumber example description.

#* Echo back the input
#* @get /custom_df
custom_df = function(num_rows) {
  
  library(tidyverse)
  
  mydf = data.frame("DataA" = 1:num_rows, "DataB" = 1:num_rows)
  
  return(as.data.frame(mydf))
  
}