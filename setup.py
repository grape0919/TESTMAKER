# -*- coding: utf-8 -*-
from setuptools import setup

# name, description, version등의 정보는 일반적인 setup.py와 같습니다.
setup(
      # 설치시 의존성 추가
      setup_requires=["py2app"],
      app=["Program_main.py"],
      options={
          "py2app": {
              "includes": ["pandas", "PyQt5"]
          }
      })