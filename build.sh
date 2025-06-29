#!/bin/bash

# initialize the submodule
git submodule update --init --recursive

# install the requirements
pip install -r requirements.txt

# start the app
python app.py
