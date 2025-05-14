#!/bin/bash




read -p "Do you want to disable or enable a devices(D,E)" option

read -p "What is the asset tag of the laptops you want to disable or enable? " asset


if [[ "$option" == "D" ]]; then
  gam update cros $asset action disable
  echo "$asset has been disabled"
else
  gam update cros $asset action reenable
  echo "$asset has been reenabled"
fi
