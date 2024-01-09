# proc_teachingschedule

## MacOS

Change default shell to bash:

chsh -s /bin/bash

Setup bash

Add a .bashrc file with...
# Source global definitions
if [ -f /etc/bashrc ]; then
  . /etc/bashrc
fi

# Uncomment the following line if you don't like systemctl's auto-paging feature:
# export SYSTEMD_PAGER=

# User specific aliases and functions
#export PS1="\e[0;35m(^: \033[1;33m\033[49m\e]0;GO UW HUSKIES\a"
#LS_COLORS="di=0;1:ex=0;32"
#asdfasd


Then add a .bash_profile

if [ -r ~/.bashrc ]; then
   source ~/.bashrc
fi




Install IPython3 

pip3 install --upgrade pip
pip3 install ipython 
