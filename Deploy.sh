#!/bin/bash

ROUTEVAR="/mnt/c/Users/dsche/OneDrive/Documents/Programming Projects/Google Sheet Code/Financal Tracker/code"
npx tsc && ./MovetoRoot.sh > /dev/null && cd "$ROUTEVAR/out" && npx clasp push
