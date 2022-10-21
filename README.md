# LV_cavern_mapping
Repo contains mapping files (which may need validation/correction) as well as scripts that re-format the mappings and check for consistency.

In particular, given a (specifically formatted) cavern mapping (stored in `formatted_cavern`), running

`python parseXls.py <formatted_cavern_mapping>`

will produce cavern mappings (formatted by Mark) in a convenient form for the software (and potentially also for checking the cabling in the cavern) in the `output` folder.

The cavern mapping was found to have mistakes because the PPP mapping had changed and the cavern mapping wasn't updated (or, potentially it was just copied incorrectly). Running

`python check_fix_mapping_ppp.py <mapping_with_nominal_PPP> <formatted_cavern_mapping> (<check_against_compare_mappings>)`

will first check for (not easily fixable) typos in the formatted cavern mapping file (by making sure that all the nominal lines, which are trusted to be typo-free, can be found in the cavern mapping), outputting lines with typos to the command line. Then, once any typos in the cavern mapping are fixed (by the user), running this will check the nominal mapping vs the formatted cavern mapping (and, optionally, the mappings included in the compare file, if optional third command line arg is `true`--NOTE this isn't implemented in the code right now; still, include a third argument, say `false`) and output PPP mapping mistakes to a csv stored in the fixme folder.

Not implemented here: add cavern mapping with extra column indicating the correct PPP info in the unformatted_fixed_cavern folder. To then produce the software-usable cavern mapping, the user should take the unformatted cavern file, format it however desired, delete the old PPP info columns, store this edited file in the formatted_cavern folder, and run `parseXls` with it as input.

Note (not necessary, not sure if would work with current implementation anyway): if desired, it would be very easy to accommodate for any alteration to the PPP mapping; `check_fix_mapping_ppp` would just have to be run with a sheet formatted like the current surface mapping (but with updated PPP info), and then following the same workflow as above.
