# Basel Tools

`Basel Tools` is a set of `GUI` applications that can produce a rough estimation of capital requirements based on the `Basel IV` regulatory framework proposed by the `Basel Committee on Banking Supervision`. The following applications are currently available:
* `BaselCCR`: a tool for calculating the capital requirements related to the counterparty credit risk exposures.
* `BaselOP`: a tool for calculating the capital requirements related to the operational risk.

## BaselCCR

The tool can be run by executing the `BaselCCR.m` script. The underlying calculations are based on the `SA-CCR` model defined within the `BCBS 279` document (https://www.bis.org/publ/bcbs279.htm), but the application also offers the option to compute the counterparty credit risk exposures using the `Simplified SA-CCR` as per `CRR II` european regulatory framework.

## BaselOP

The tool can be run by executing the `BaselOP.m` script. The underlying calculations are based on the `SMA` model defined within the `BCBS 356` document (https://www.bis.org/bcbs/publ/d356.htm). The application offers the opportunity to compare the `SMA` capital requirements with those produced by the obsolete `Basel II` approaches defined in `BCBS 128` (https://www.bis.org/publ/bcbs128.htm): the `Basic Indicator Approach`, the `Standardised Approach` and the `Alternative Standardised Approach`.

## Requirements

The minimum Matlab version required is `R2017a`. In addition, the following products and toolboxes must be installed in order to properly execute the script:
* Financial Toolbox
* Statistics and Machine Learning Toolbox

## Notes

* Some core functionalities of all the `Basel Tools` applications are defined within the `BaselTools.jar` package, which has been compiled under `Java 1.8` and includes the original `.java` source code files. Depending on the current Matlab version being used (the console command `version -java` should provide enough details), it may be necessary to recompile it referencing the proper `Java Framework` version.
* In order to have a full control over the underlying `Java` components, all the `Basel Tools` applications use the `findjobj.m` script intensively. The script, created by Yair Altman, can be found by visiting [this page](https://it.mathworks.com/matlabcentral/fileexchange/14317-findjobj-find-java-handles-of-matlab-graphic-objects) and is always included into every `Basel Tools` release.
