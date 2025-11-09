#!/usr/bin/env bash

# Script to render all Quarto (.qmd) files in each subdirectory

for dir in */; do
    if [ -d "$dir" ]; then
        cd "$dir" || exit
        for file in [0-9]*_*.qmd; do
            if [ -f "$file" ]; then
                echo "Rendering $file..."
                quarto render "$file" --to html
            fi
        done
        cd .. || exit
    fi
done

echo "All Quarto files have been rendered."