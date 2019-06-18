#!/bin/bash

cp .clasp.json.prod .clasp.json
clasp push
cp .clasp.json.dev .clasp.json