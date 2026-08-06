@echo off
cd /d C:\Users\Usuario\conversor markdown
git config user.email "usuario@local"
git config user.name "Usuario Local"
git add -A
git commit -m "Marca: Preparador de archivos para IA - arreglos de build e icono"
git rev-parse --short HEAD
