#!/usr/bin/env python3
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.write("script start\n")
sys.stdout.flush()
import md2docx
sys.stdout.write("module loaded\n")
sys.stdout.flush()
md2docx.parse_md_to_docx("test_simple.md", "test_simple.docx")
sys.stdout.write("conversion done\n")
sys.stdout.flush()
