###################################################################################################
# Create new environment for use by package functions

acctabs <- base::new.env(parent = base::emptyenv())

utils::globalVariables(".")

###################################################################################################