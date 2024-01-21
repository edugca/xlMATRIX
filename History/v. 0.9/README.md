# xlMATRIX (v. 0.9)

Handle vectors and matrices in Excel as easy as in MatLab: slice, reverse, stack, shelve and more with Lambda functions. Conduct statistical and econometric analysis. No VBA! Many functions from MatLab are coded in Excel keeping their original names.

## History

### v. 0.9
- Added functions JACKKNIFE, CORREL_DIFF_TEST, PCA, COV, MAHAL, OLS_VARCOV_HC, OLS_COEFF_STDEV_HC, OLS_COEFF_TSTAT_HC, and OLS_COEFF_PVALUE_HC.
- Functions OLS_VARCOV_HAC, OLS_COEFF_STDEV_HAC, OLS_COEFF_TSTAT_HAC, and OLS_COEFF_PVALUE_HAC have two new optional parameters [type] and [lags]. Previously, the amount of lags for the HAC calculation was fixed by a formula, which remains applicable when the argument is omitted. The type can be "Bartlett" (default) or "Parzen".
- Function OLS_SUMMARY has new optional parameters [robustStd], [robustType], and [lagsHAC]. The parameter [useHAC] was abandoned.
- Function EIGENVALUES and EIGENVECTORS now have 0 as default value of the parameter [order]. This will present results without sorting them.

### v. 0.8
- Added functions CORRCOEF, VARINFLATIONFACTOR, ISMULTICOLLINEAR, CONTAGION_FORBESRIGOBON, CONTAGION_COVOLATILITY, CONTAGION_COKURTOSIS, CONTAGION_COSKEWNESS, FACTORIZE_QR, FACTORIZE_CHOL, LINSOLVE, SUBS, OLS_HQC, OLS_SUMMARY, OLS_RESIDUALS_AR, HASCONST, VERTCAT_EX, HORZCAT_EX, BOOTSTRP, RANKRANGE.
- Functions EIGENVALUES and EIGENVECTORS now are calculated through QR decomposition (much faster than previous algorithm based on finding roots of a polynomial).
- Description of functions EIGENVECTORS and MNORM were updated.
- Functions LAG and LEAD contained reference to function REVERSECOLS, old name of function FLIPLR. Now, they accept a matrix of column-vectors as input.
- Functions FLIPLR and FLIPUD have a new algorithm and a new optional parameter [nSize].
- Function OLS_LOGLIKELIHOOD has a new formula to match statistical packages. This change also impacts results of OLS_AIC and OLS_BIC.
- Function ADDTREND now has a parameter for header's specification.
- Function ADDCONST has a new algorithm, but should keep the same functionality.
- Function REPMAT now has a parameter, [expandFromWithin], for expanding arrays by repeating each element on the directions requested.
- Functions OLS_ESS, OLS_RSS, OLS_FSTAT, and OLS_FSTAT_PVALUE had their algorithms mended for regressions without an intercept.
- Functions ALL and ANY had their algorithm changed to accept arrays as input instead of only arrays of cells.
- Function SAMPLE had its parameter "replacement" renamed to "withoutReplacement" to avoid confusion.

### v. 0.7
- Added functions MOVAVG, MOVSUM, MOVPROD, SEASFACTOR, TSTREND, STARTSWITH, ENDSWITH, CONTAINS, ISMEMBER, AGGREGBYDATE, EIGENVALUES, EIGENVECTORS, MNORM, EVALS, FZERO.
- Function SEASADJ's description was updated.
- Function DETREND now allows for log-detrending. Parameter "nPower" was renamed to "type".

### v. 0.6
- Added functions GET, ISOUTLIER, RMOUTLIERS, ISEMPTY, MULTIFILTER, TEXTTOARRAY, REPLACEALL, SPLITTWO, SPLITALL, COMPLEMENT, DIFF, DFTABLE, ISSTATIONARY, STATIONARIZE, IORDER, ISCOINTEGRATED, TSOPER, OLS_COEFF_TSTAT, OLS_COEFF_TSTAT_HAC, OLS_COEFF_PVALUE, OLS_COEFF_PVALUE_HAC, OLS_STDEV, OLS_STDEV_HAC, OLS_ESS, OLS_RSS, OLS_TSS, OLS_FSTAT, OLS_FSTAT_PVALUE, OLS_R2, OLS_R2_ADJUSTED, OLS_LOGLIKELIHOOD, OLS_AIC, OLS_BIC, OLS_OPTIMAL_LAG, OLS_LEVERAGE, OLS_COOKS_DISTANCE, OLS_RESIDUALS_STD, OLS_RESIDUALS_PARTIAL, OLS_CONF_INTERVAL, and OLS_PRED_INTERVAL.
- Functions SLICE and M now accept matching texts for the matrix coordinates. That is, if instead of a number for lines and columns the user provides a text, then the function will search for that text in the first row/column and set its position as the coordinate.
- Functions SLICE and M have a new parameter. The parameter "patch", which receives a matrix, induces the function to return the original matrix patched by the "patch" matrix in the address specified.
- Function SLICE has a new parameter. The parameter "remove", which is a boolean that changes the behavior of the function so that indicated lines and columns are actually removed from the original matrix.
- Function M now accepts references to negative positions (start counting from the last element) and also interprets "0" in the address as the first or the last element depending on whether it preceeds or not the ":".
- Function SPLIT now accepts delimeters with more than one character.
- Function NOZERO returns #N/A when there are no filtered results. In MatLab, the equivalent function returns an empty array.
- Function NOZERO has na additional parameter returnRowAndCol, which is a boolean that when TRUE the function returns two columns: row and column index.
- Functions HORZCAT and VERTCAT now return the non-omitted matrix in case either matrixA or matrixB is omitted. Besides, concatenation of matrices with different number of rows or columns is done by filling exceeding cells with #REF! or #N/A.
- Fixed bug in function ADDTREND. When parameter nPower was missing, the function returned a vector of ones.
- Function ANOVA code was slightly adjusted to this version changes to HORZCAT.
- Function ADDCONST has a new parameter. The parameter "header" allows for the user to provide a header name for the column with the constants.
- Functions LAG and LEAD have a new parameter. The parameter "headerFirstRow" is a boolean which tells the function to ignore first row values of the input and add a header row to the output matrix.
- Functions OLS_ that allow for coefficients to be provided by the user now accept that these coefficients are passed as a row or as a column vector.
- Function SAMPLE had a parameter ("replacement") with misleading behavior. It was fixed.
- Function SIZE was renamed to DIM due to the fact that SIZE is a reserved word in Excel when English is set as language.

### v. 0.5
- Added functions LAG, LEAD, GRANGER, and ANOVA.

### v. 0.4
- Added functions DETREND, SIZE, and HPFILTER.

### v. 0.3.1
- Added parameter nPower to function ADDTREND.

### v. 0.3
- Added functions ADDCONST, ADDTREND, DUMMYVAR, and SEASADJ.

### v. 0.2
- Functions REVERSEROWS and REVERSECOLS were renamed to FLIPUD and FLIPLR for consistency with MatLab names.

- Functions STACK and SHELVE were renamed to VERTCAT and HORZCAT for consistency with MatLab names.

- Added functions M, SLICE, VEC, SAMPLE, NOZERO, EYE, ZEROS, ONES, DIAG, TRIL, TRIU, FLIPLR, FLIPUD, REPMAT, RESHAPE, BLKDIAG, ALL, ANY, NNZ, LINSPACE, LENGTH, NUMEL, and RMMISSING.

- Added functions SPLIT and TEXTTOCHAR.

- Added functions OLS_COEFF, OLS_COEFF_STDEV, OLS_VARCOV, OLS_FITTED, and OLS_RESIDUALS.

### v. 0.1
- First release.










