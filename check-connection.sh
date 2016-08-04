#!/bin/bash

wget -q -t 2 -T 5 --waitretry=5 --spider http://google.com

if [ $? -eq 0 ]; then
    echo "Online"
else
    echo "Offline"
fi

