# Парсер сайта Lenovo

Цель: выгрузить 12000 спецификаций оборудования с сайта 
Результат: все получилось, но далеко не с первого раза 

Основная библиотека: Playwright
Применял параллельные вычисления, браузер Chrome for Developers, пришлось разбираться с сессиями / cookies, запускал код на виртуальный машине Linux

Практическая направленность: выгрузить все 12000 спецификаций автоматически, чтобы ускорить время 
Вручную это пришлось быть делать месяц
Код сделал это примерно за 24 часа суммарно

### How to start

Copy a folder to the VM
`orb push ~/Downloads/parsertest`

Inside a VM
`sudo apt install tmux -y` (meaning Yes to all the prompts)
`tmux`
`xvfb-run -a uv run python manualapp.py`

Detach
`Ctrl + b`
Release both 
Press
`d`

Verify
`ps aux | grep python`

Then to get back to the process
`tmux ls`
`tmux attach -t <NUMBER>`

In the tmux view just run
`Ctrl + C`
to stop the process

`tmux kill-session -t 0`

### Running a browser in a VM

Search: "How to run a browser in a VM"
https://forums.linuxmint.com/viewtopic.php?t=445104

Copilot session
https://github.com/copilot/share/0a07429a-0924-8495-a843-b20d204b6184