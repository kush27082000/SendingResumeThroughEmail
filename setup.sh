#!/bin/bash

if [ -d "./Personal_Data" ]; then
    echo "Setup folder already exist so won't overwrite. Please delete the folder ./Personal_Data if you want to re-setup"
else
    mkdir Personal_Data

    cd Personal_Data

    echo "Creating nessasary files in ./Personal_Data"
    echo "Please overwrite them without changing filenames"

    # Gmail account setting
    printf "userName = 'YourUserName@gmail.com'\n" >  gmail_account.txt
    printf "passWord = 'YourPassWord'" >>  gmail_account.txt
    printf "realName = 'YourRealName'" >>  gmail_account.txt
    printf "test_mode = True" >> gmail_account.txt

    # CV template
    touch CL_1.html
    touch CL_2.html
    touch CL_3.html

    # Personal Resume
    touch Resume.pdf

    # Transcript is an option
    touch Transcript.pdf

    # GRE score is an option
    touch GRE.pdf

    echo "Done!"

fi

wget https://pypi.python.org/packages/source/x/xlrd/xlrd-0.9.2.tar.gz
wget https://pypi.python.org/packages/source/x/xlwt/xlwt-0.7.5.tar.gz
wget https://pypi.python.org/packages/source/x/xlutils/xlutils-1.7.0.tar.gz

gunzip -c xlrd-0.9.2.tar.gz | tar xopf -
gunzip -c xlwt-0.7.5.tar.gz | tar xopf -
gunzip -c xlutils-1.7.0.tar.gz | tar xopf -
cd xlrd-0.9.2
sudo python setup.py install
cd ..
cd xlwt-0.7.5
sudo python setup.py install
cd ..
cd xlutils-1.7.0
sudo python setup.py install
cd ..
rm -rf x*
