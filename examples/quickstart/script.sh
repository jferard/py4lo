#!/bin/bash -x

mkdir -p qs/src/main
cd qs || exit

cat > src/main/qs.py<< EOF
# -*- coding: utf-8 -*-
# py4lo: entry
# py4lo: embed lib py4lo_typing
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_dialogs
from py4lo_dialogs import message_box


def test(*args):
    message_box("A message", "py4lo")
EOF

python3 ../../../py4lo init
mv new-project.ods qs.ods

cat >qs.toml<< EOF
calc_exe = "libreoffice"
python_exe = "python3"

[src]
source_ods_file = "{project}/qs.ods"
EOF

cat >qs.py<< EOF
# -*- coding: utf-8 -*-
# py4lo: entry
# py4lo: embed lib py4lo_typing
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_dialogs
from py4lo_dialogs import message_box


def test(*args):
    message_box("Another message", "py4lo")
EOF

python3 ../../../py4lo -t qs.toml run
