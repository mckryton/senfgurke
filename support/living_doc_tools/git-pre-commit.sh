#!/bin/sh

#
# update and include the features index as part of the commit process, but only
# if needed.
#
# This is not exactly 'best practice', but it should prevent a few lazy
# slip-ups
#

# installation:
#
#   cd <project root>
#   ln -s ../../support/git-pre-commit.sh .git/hooks/pre-commit
#

feature_index="features/README.md"
need_update=$( git status --porcelain 'features/**.feature' | wc -l | tr -d ' ' )
if [[ $need_update != 0 ]]
then
    echo "Updating $feature_index"
    ./support/living_doc_tools/generate_index.pl --index $feature_index
    git add $feature_index
fi

exit
