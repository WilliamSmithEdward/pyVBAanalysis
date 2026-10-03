#!/bin/sh
# Run the six corpus parts (all.0.json .. all.5.json in the current folder)
# through one XLIDE tree in parallel, into all_TAG.jsonl.
#     sh tools/hunt/sweep6.sh TREE TAG
here=$(dirname "$0")
tree=$1
tag=$2
for i in 0 1 2 3 4 5; do
  XLIDE_ROOT=$tree npx -y tsx "$here/wrapper_probe.mjs" all.$i.json > all_$tag.$i.jsonl 2> all_$tag.$i.err &
done
wait
cat all_$tag.0.jsonl all_$tag.1.jsonl all_$tag.2.jsonl all_$tag.3.jsonl all_$tag.4.jsonl all_$tag.5.jsonl > all_$tag.jsonl
wc -l all_$tag.jsonl
