#!/bin/bash
cd /home/kavia/workspace/code-generation/presentation-creator-7036-7045/ppt_generator_frontend
npm run build
EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ]; then
   exit 1
fi

