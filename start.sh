#!/bin/bash

source ./venv/bin/activate

echo -e "4 xls files into csv for shop.\n You have to write tables in specific order:\n avtozavod, svoboda, varya, telegraphnaya"

echo "Avtozavod's excell file location: "
read avtozavod

echo "Svoboda's excell file location: "
read svoboda

echo "Varya's excell file location: "
read varya

echo "Telegraphnaya's excell file location: "
read telegraphya

echo "Output folder"
read output_dir

python3 main.py $avtozavod $svoboda $varya $telegraphnaya -o $output_dir

deactivate