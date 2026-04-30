#!/bin/bash
set -e

aws ecr get-login-password --region ap-southeast-1 | docker login --username AWS --password-stdin 510241986114.dkr.ecr.ap-southeast-1.amazonaws.com

docker build -t 510241986114.dkr.ecr.ap-southeast-1.amazonaws.com/trainocate/prompt-checker:latest .

docker push 510241986114.dkr.ecr.ap-southeast-1.amazonaws.com/trainocate/prompt-checker:latest