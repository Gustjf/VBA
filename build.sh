#!/usr/bin/env bash
# Build LaTeX sources into the mirrored output tree.
#   source/<Topic>/x.tex  ->  output/<Topic>/x.pdf
# Usage: ./build.sh [file.tex ...]   (no arguments = build every .tex under source/)
# Needs latexmk (sudo apt install latexmk texlive-latex-extra texlive-fonts-recommended).
# Aux files go to .build/ (git-ignored), so output/ only receives the PDF.
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$root"

if [ "$#" -eq 0 ]; then
  files=()
  while IFS= read -r -d '' f; do files+=("$f"); done < <(find source -name '*.tex' -print0 | sort -z)
else
  files=("$@")
fi

status=0
for f in "${files[@]}"; do
  abs="$(realpath "$f")"
  rel="${abs#"$root"/}"
  case "$rel" in
    source/*) ;;
    *) echo "skip (not under source/): $f" >&2; continue ;;
  esac
  dir="$(dirname "${rel#source/}")"
  out="$root/output/$dir"
  aux="$root/.build/$dir"
  mkdir -p "$out" "$aux"
  echo "==> $rel -> output/$dir/$(basename "${rel%.tex}").pdf"
  if ! latexmk -pdf -cd -interaction=nonstopmode -halt-on-error -file-line-error \
        -outdir="$out" -auxdir="$aux" "$abs"; then
    echo "FAILED: $rel (log: .build/$dir/$(basename "${rel%.tex}").log)" >&2
    status=1
  fi
done
exit "$status"
