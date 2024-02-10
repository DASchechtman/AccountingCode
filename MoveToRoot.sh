#!/bin/bash

target="./out"
find "$target" -mindepth 2 -type f -print -exec mv {} "$target" \;
find "$target" -type d -empty -delete