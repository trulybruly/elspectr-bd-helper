#!/bin/bash

source ./venv/bin/activate

echo -e "4 xls files into csv for shop.\n You have to write tables in specific order:\n avtozavod, svoboda, varya, telegraphnaya"

echo "\nAvtozavod's excell file location: "
read avtozavod

echo "\nSvoboda's excell file location: "
read svoboda

echo "\nVarya's excell file location: "
read varya

echo "\nTelegraphnaya's excell file location: "
read telegraphya

echo "\nOutput folder:"
read output_dir

python3 main.py $avtozavod $svoboda $varya $telegraphnaya -o $output_dir

deactivate