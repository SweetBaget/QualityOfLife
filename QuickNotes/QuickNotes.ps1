#requires -Version 5.1
<#
.SYNOPSIS
    Быстрые заметки (QuickNotes) — лёгкая утилита для Windows: создание текстовых
    файлов в один клик.

.DESCRIPTION
    Каждая пользовательская кнопка привязана к своей папке. Нажатие на кнопку
    открывает окно заметки, в котором имя файла уже заполнено текущей датой,
    а текст можно отредактировать и сохранить в .txt.

    Кнопки хранятся в файле buttons.json рядом со скриптом (если папка доступна
    для записи), иначе — в %APPDATA%\QuickNotes\buttons.json.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File QuickNotes.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File QuickNotes.ps1 -SelfTest
    Запуск встроенных проверок (без GUI).
#>

[CmdletBinding()]
param(
    # Запустить встроенные проверки логики и разметки, не открывая окно.
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'

# ============================== НАСТРОЙКИ ====================================

$script:AppTitle   = 'Быстрые заметки'
$script:Extension  = '.txt'
$script:DateFormat = 'yyyy-MM-dd'      # формат имени новой заметки по умолчанию
$script:ConfigName = 'buttons.json'

# ============================ ИКОНКА ПРИЛОЖЕНИЯ ==============================
# Иконка встроена в скрипт (base64), поэтому она есть даже при запуске одного
# .ps1: окно, панель задач и Alt+Tab. Файл AppIcon.ico рядом со скриптом —
# резервный источник иконки, им же назначается значок ярлыку на рабочем столе.

$script:Base64Icon = @'
AAABAAcAEBAAAAAAIAB0AgAAdgAAABgYAAAAACAA3AMAAOoCAAAgIAAAAAAgAFwFAADGBgAAMDAAAAAAIACJCAAAIgwAAEBAAAAAACAADgsAAKsUAACAgAAA
AAAgAL4VAAC5HwAAAAAAAAAAIAD4KgAAdzUAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAAjtJREFUeJxtk7uLFlkQxX9VfbuHee18opMo
CBptYq4bGKmRmamRgQqGZoNgoJjoJgsLjqn+FbuaCcYmgokggojKzOjMp3bfe+sY9DcPxBsWdQ7nUdcA7t+eXm/bbi0P+WioOsIkIUFIKIQkIiTJw719n/tv
9+78s/rQHtyeXu26hfWcB0oZKAVKGQG7QAkihDuYB24t7h0/hp1rSeJuzlm59IpqvrQMyytOVCFGsEKYwdZm5csWVPpwl1F1NwlWS+lBZpK4dHmeY8cTv3vv
3hbW/94CzGvpEbaaIkIIixBNAwtLRt/3bG/3uBsAZs5kssjiotEkyIMAiAil/cBEKUHbOk+fPef//16wvLzAMGQmkxVu3rxCKFHruAsgYSkkEESAu9H/CC6c
P8tfZ07PFAhwUkpIAeIAgUiaDSQREmbOzs6UjY0pTdNgBqVUaizitkQoiADQjGCvKu2xt6lhbq7F3QFoGqdtE33RAQVjzWkXLAmCWWgGGGa7+dvohNlh7Wew
SwAKkIkmiTev53n1co6UYOWQce6iYwbfp2XcmxGERIqQAAuJxqEG/HkKjp803MEdch7baZqZ9/0MxhpjnJAH8W07OHy4oev29BMBZjDdCYZBGJpZGWv85NYe
ybVXhPmTR1/5Y+JEHPAaAoONz4UIYUSYt1br8Dmpxi3vunU3maxnaxM+fSy/fCYRAU0D5sJo3a1DMdwygLUbH66ntLBW6nA0ojiM1zmqOHAnVRIeZu37nKf3
/n184uFPZ1mx7R9YDf4AAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAGAAAABgIBgAAAOB3PfgAAAOjSURBVHicdZXBixxFFMZ/r6qmZyea5KDuBkHJ
Qo6GYBZkReMhCgrG4/wh7sGbhiV7MuJqQP8B9RTwJiaKCEtANuiKl5wlIOghGt3s7HT31HseqqunezZbMFP9uqq+9973vnot47H5mzclXn//0duDwfC9qLO1
qPGEGWJmYGBmpMc0k23r2GAifiL4X+rZ5MNrnyx/Mx6bF4DrHxy8OxwufawqVPVh/2AG4ljgxjYwwfslQKiq/Y2tG8vb8tHm4ZtFWLpVVYcWNSqINwOhD3Zs
9M17VQMz1CwKzoUwknK6/1Yw1auq2gFPtJSlYRiSMBsHc/BMnWqaB4U0DsVHjdGpOkSuBjO7WNWlZHABysp44cXA2nqBplOdYYC0TsTB3Tslv96dJidqAL6u
DzG1iwFkaKZN+qCAc/Dq6wXPrHjUFJGuE2n+HeAAOPGE8NvP0wTe4JgpBsPwuIKKwGyWqHHiOG5UFQxC2iuOlG23bmaEnhQz340TQ/n+u5948OAfQvDNeyHG
yKlTJ7l8+RXEhbS3xZn/AEJf5/NiioCqsrOzy/37f1AUReugrmuWl5/m0qWXWFoKWQXzIJmrK/SUQbuaFn1ga2vjWIrmZZGObPt0h1aC0OGvzYW//nxIjIpI
H1PVCCFw5szpNqAcZPeehHTJu14b/kzAYP+/CVVdIwsezIyiGLCycrq1WxI6wgl2BLy5UCS+z64+O+ftKDfJcW4l9HEwCF2+cg3MQBpe7937nem0xDlp18Wl
/cPhkPPnz6YaaKeGuejkGnTTazxkxVy48Bz59uZR10ZdpyByIJma1seiiuadEdQSnCp8/ZXx8G8jhHS4ruCNdwKr5xx13a/JYg1aFTVxd7hLtgisnhMmB+B8
OhwjPHkSYsyA0iqn375TKqGXXucG5nntZX+ktKqpPeTn7pkcaJZ86KXXZBOjETxNuxAWrgDOQVHM7RCEGC01xQ64kTIosdRRcy1nM+OHbyesvzaaq2thZDk6
gTs/HjCrlaIQtFMLVSuDme0Nwmi9rA7SB0eNwUDY252ytzsla7v3JetQqmqoWgLXdk8MfuSiTvaCqWyayS3BOdUYDTwGxTBp20wa6eZ+kz82jd0QqLEVSQRx
hoiZbLprnz51u6r2N0IYSQgjnw9qNDRHrDlSGqep96slO4E3nPuR934kZbW/8dkXz99247H5rRvL22X57xWNuiO4g3TPrE/LIk2LMxi4A9W4My0fXfn8y9Xt
8dj8/yELhHxd5SK7AAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFI0lEQVR4nJ1XTYgcRRT+XnV198wmu/kh0d3EuCxx
JRBUPAoeREECnjcHj540uSSHiIiQKLIENkiIJuaq3nbxJhKiF/EqSMBoYiKSCPnDU2Zmf7q73uehu6Z/pmdXUoeZ6npd9b3v/dVrAYCFBQYrK+LOftQ7HMfx
CVV9U53uI9USAAgw/wFrcwJoruUbyFyQP0omYu5BzNU065//9PzMdY8py8sMjh4Vt3R6cCy04dkgCCeTJIHTLAcYB14gVwFr4MgV9GsiFqHtIsuS3kY6+GDx
wtOXlhcYCAAsnR68O7l94svBYB2Zcw6AASBPDF68W5eTpKoRG3Q7u9DrPXxv8eL0ZTl3+vEhMdGvACKXZYCIaTIXaQdvNXuFuWpdIRBQVTXGgkTi1L1syeBU
HMad1bW+EzFBm9nX1xSqW4NXY0KEiGJBIS3PFDFOUxdHUx238fiUBXgkSRICo8y1CKRDL4SY2C6gFhRrbDEcfj8EGPQVf/6egiRkGKi5NUCYNF0jgCOWxLTT
zKs6BIcAyZrijbdivPp6jCcZP/2whu+/HaDTFWiVHEScy6DktAUwwpwAoERgBQeftyCBJCFE2I7klyX/IYEoEswfChGEeSw03QoQAjGtee71BIgsI0RynwaB
gch4xqrVGACyrCQ0LmbsuFSrHgQAYWjgHJEkGaRFC5LodEKQBTCK7CniaFzA2q3yXIsXb926i8XFz1uZiwjSNMX09FM4c+YEup1OxS8sA7A2zzHtOPC6c4Eo
CrFnzy44p4UFSlmuQIbdu3fCGIPqbjYI1WsGYbeqcN7cs7MzOHfuw5pb2oZzgHNFGnk1PMkGONlwQU1Ys0TuV2Ok9GtleKVIQkRgjNRkHNqrDo6hC9qEQxPm
D6ur67h9+59N2ecpSDx38AB27OyWhYmjlq24oEydWpAUG6QwZRAEmJrathn6cATWFOoIiFHf11wwAu4LQykAAExMRJif3/+/FFAt5+UR7Ve23fQ+R1mS+v11
3LlzfwsX5GfMzs5gaqpTqS9tV3Z+8pggRN1//nAFOEYBk1u9rmCRYU0reHJlELaBF+x9Gk5OdfDiS3Ob0y9GlhWTSsaU3VFJjr4OtDH3mvsb+N+Hip9/zFAd
IkCaEPsOGLzymi1872uA+AMr5h8931ZzstpG+c1+vrZK/HVTayYWAZINn/9+Teo9QkXhJjg4DMI6eNVPHu+ZWYNTn4zvC/JurrHoibEdnCCsj8YmeP0GKwHa
smBcea7dri3guQWa4F5YmMxaNHzYboGa2RWgBWwolVI8Ck4AFkrNO+G6UARIU+LWHyn2PxsijrcoAC3j5m/rSBOFnTB5gDbAqapWgQeBBDOZpoD/Fih8FkWC
q98NcPfvFNu2m2FvgKrFysfhmgjQ7ymuX1tHFAvUsQ5OUowFwQcWwBVru++k2YYTmIDeAQQghBC49svGsMdnoSEpBeDoDUcQAkHcRaUjrjCHamg6wYbrXbEi
urSRDN42EkXOJSoipmmqzgQAylAxFuUwB5ZCCb9G+MZUHUfBVVUklCRdXYfokvn4s703Mped7HZ2GGMCo1RHqicKkFCXNxrqmP8r4Yq5G67lcx2uNT9gSKo6
kcB04p1GmZ784qu5G2Z5gcHihb2Xe/1Hx41EvTicCowJpX5JseLzMRfLmDz3zI1YicLJQMT2VgePjl/8Zu7ygv849Z/K7x+7fzgOu/nnuWb7SLb2C/WqOb7I
eHAQmUhwDyJXU9c/f+nr+esLC8vByspR9x+JEgCkOvKutAAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAAAwAAAAMAgGAAAAVwL5hwAACFBJREFUeJzF
mk2IHMcVx/+vqrp6VquVFEtBkkVkmQSsGCwHhVhsbOT1yQh8MuwiYvDFxIfgFT4F4nwsm4MSTL5wNiGXYBAhCOlmdDE+SMhxNtElgZxEZMdWtJb8lRB5Z2em
u+u9HLp7uvprZrRa7Hfoma6uj1+9eh/d1U3w5Nw50QsL5ADgVz+OnxLISSc8K84dEMBChCSvLICkh/wUEMl+szIR71p6kGEH3nUZFnjdUURQa0S06ljOLv9i
5wUAODcveuF8yggAVIV/+Ye3H7XB1GkQHVdKI0liOE4gLPC7F29kGZZNNqGiafuEiABFBkqHYBdDWC73o42XfrKy9y1/EgoAlpYumoUFci//qLsYhlNvamOO
R3FP+oOui10kwvyZwgOAsCBxkUSDT10cb4g29ngnnH7z+4sfLC6cJ7e0dNEAAOWa/+nSpy/s2Lb91xu9HjM7IShdAvkM4etjCoTZgTRNdXapbvfDxdMr+1fO
zYsmAPj5UvcRE9i/xknimBNFUCSVzj9P+HwIYRalNCttdRwNjp1e2XOFlpZEbcfGnzrhttlef90NNf+5w4s3TlGJhV3H7tT9we3VYPcXHzPbefBkYMPZfr/L
k8ITACgBZScCgEQghKJMBEQFPJEMnTPvn6hwWGEBlzTv/RbFICjdG/yPA9OZjT/56EkjlDyrVCh+3XHwUSyIBwVArtmS9iZZBW8VTQAElsBcGbthNQgQRYEw
+s8aALNJkpAAimqd1+EHA8a+AxoPHjEwQa5ClFSV182BfZHhxaJeHAn+8bcIa9cT2BAQbofPzlXiBgTBrCGh/Y5jUGYRIzUfMb50n8Yz394GGw5TyJbI7OMd
vPqb23j37RjWUjp0ix8AROwSANhvhGDBPDZUQgmSBPjGYxY2JPR7AqW3Bp4d0JkizM51cO1qDGvb4YsyBgTWlMPCiGjDgFJA2Em9Vun0fMtEgE6HoFTqN+3w
5cilpATaAj90yDRaYGutJ5U8IIzRvA8vAAxGVG6K81UREag7XApmBlGzFkbD16ObmRSeGuABIAgUoshlHVbCUYuEoYFz5RBbnkGVp2w2fjtT13KL5qm8AKnm
CWfPXsDFi6sIQ9sM5AkRIYpiPPzwg3juuQWohiiQmvQIs6nkFTMRPFCxHoHWClEU47XX3sD773+ITmf8BACCcw7Xr6/h6adPYO/eLyCOGblTSYYPoWaz8Qmy
EzMOPr9GVGgmBWGEYYDnn/8WLl36C4IgmHAFIhw9+hD27NmFOE59wU+ehUlX4FsyupkEfqgdj4+I4Jxgbu4Y5uaOjQRvkiSRuiPTpPD+CtRAG+BFhpGzquPB
wKEloLSKCKAUAaByW6E6fKVddUJmEvhqB75Ye3fp2LkSItK1ptaxxzjx+Pv5shkBN258hNu3u5vKBTt2TOPee/eUysX3g5HweR6YED6/n/fhmQW3bn2C9fWN
TU2g2+1h377dNV8YF/t9pzf1rY8WzVdEBNCacPjwIXS7PSilJgijqRARmBnT01NQiuBccW08fGEygkoYbYOv5gD/+vR0BzMznYnAq8KM4gGmNCiN1Xz+19wJ
fFXDuRn1+8kmIpHA2qCx3ajY79/6pCuwWfis3tWr723aiWdmtuHw4UPFGK3wZbMZ+isAUwdths+ff6tWLv663q00+FvNbDz4zInHwzfxCQBFwAMPHEIUJWPZ
iFAzlyAwtTKRarb1zMbH8J14HHxxp1xeAhEgCIAgCMZOAEAp2uT9twWuNpsv4NMLZiL4rFU1uSgF3FoT3LzhoA1aYZiBg/cr7LqHwFxfic3CC/w8MApepGpB
UApIEuD3r0T497s83A6pilJAbwP42iMa3/muBTfUKcNXHBZVsymblLcCo+GpVg/QGvjmnMY7/yQEQfMKEAFxDDx0VLeukNdr+d8IzedlZmLNt0SIJ04YPHFi
HFgqSTKZ+dTiRiN8ngcmhBdqiKEA4gjjdymyJRwLX5qIb0p1zef6ND5oM7xnQqjPgSbJX3eSpf1IN0LzeYEZDV/MNC9Tft0tFJGKksZoPm0jFSce4TQK6cNH
v59m5WpMvxtxDrAW6PcFzgkIlG611+ClBA8ABowIBNtkNqVQJgKtgT9f6uGrRyymprZ2ey6KBJffWIfWBBagjSf9SY1aIJFh8E2j7H1xEgkRUZPZQNKXD9YS
/nUtxu9+9l8c+XoHgaVUEy37WU3PEk0hOYoYf7/Sw3tvx7Adyt6IFg2qZgOIKBWQS+KbBoJVrcODSRIxBHqU0zAENgRuXE/wzrX1VAf1zrOydFZt0a16Hhhq
gffNZljMSlmVIFo1InTGueSkiJD/4qEKP/yV9E2KtbVOswNVysirJxhuWg0nnO5EsKAC3/7SXBjEnBCIzyi1657Xo2RjNQx3KGFO9wiybFKGL3UAZoFz6X2O
cwJ2kv33y+CVpeeOAXapszIjK5MJzCYbW8SFdkYlcW919/2HXlfLy8Sa6EXnYpAyEOGibQN8fmi6y57oScpv0xoqG80mfc1KCo5jQKkXl5eJ1bl50cu/3HNl
MOgudsKdmqCEhV0zvJThvbzf/iTlAUrhE83w0hAqc58RR6TFBjt1FK8vrpw5eGV+XrTKX9ufXtm7srHx8SmtQ+rYnZqFBcJuOGRVIzUtN5eV4MdqHlV4EYgT
ZgmDGa1VQL3ex6d++4cvryw9ftGcP0+u+Ngj+4Diey988GhgOqcBHFdkkLgo/diiqvkSaPtT3GTwTQ5LIGWglQVzDBa5zFH3pZU/fuWt+XnR57OPPUrZyP8K
5Aen/vMUQCeZeVbEHRARi+w99mizKcjL0Wb8hLy6AlAEqDUitSqEs6+8uv8CAPjwAPB/otApz6UZo7AAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAA
QAAAAEAIBgAAAKppcd4AAArVSURBVHic3ZtfjB1VHce/vzPnzv2z/7p0l9rddG0tpSD4gI3RKAapKCk18GI3JCYmhMQXus88mHjdhAaBF812QQ2QGKkmNMGE
2BhDZEOMBh7KE6lKSkoKK1aWKru9d/fOnPP7+TAz987MnXv3zuwWS3/JvXfmnnPm/D7z+/3O3xlChrx4TJzZ02QB4CePXh4rDQ0dJZGjLHwIjGkGDwFEEAEA
SFRQouP2QfATfkkiMyDSOUmmCSSRr5MY/x8iiSIg1QCwTKCzROpMw/HPPPHEDR8DwLFj4pwOmeJCXfAvijM7S7Y+tzK6Y3J4DqIeVo7eR0Sw1sBYAxHOUiAf
fIymEDxSN1AAIgWlHDjKBQvDsrkAluc+8v+1sLBw82rcsJGo+Em9vqRnZ8k+Wb/y7fFdY2crlfJjILXP85rcajWs77dEhIPatgIfli8ML9IFH/xaGL8lLW/N
el6DSbCvUh55bMKdOvvDuUvfnj1Ntl5f0nHmtgfU66Ln58k8+ePGI+VS+SSY0PKaBkQOgUhStV0Tbp9Ik7gKwbGwMMS6uqYBguc1j59YmFyMWNs3IHL7J3+0
9sjw0PDJ9fV1ttaClFLJi36K4GMKsTATOaiWx9SVxofHH1/cvRiFA0UHT9XXDpfd6p9837eWrSIiuh7gox9hFqUcdnTF8Vur3zxxcverx46JQ/W6qFpzZcgZ
Gn6r5JRmWt46Xy+W79QXnLAwl3RNGbNx0Wvg9tqeiYaanyd2akNzw7XKTMvbMNcrvAAgUsrzm6ZW3TmjazI3P09MP6uvjHpSOad1ecqYlgCktg0+fY0B4BNQ
SN3ALcC384mwo10y/sY/r6jVz2tfakddtzLteU3eCjypbngSASg6RtjkSthnd6pp9zECCHW6JhFJ5QvrAmClU9Gg8AAAgjJ+i93S8PSQUUc1g++nQAMpAk8k
YAbWGwJFACnq9PMUXiqCl5TlY199QyfuOQxYBsoVgiKAc8BLqD8RBKQELPdrIhyy1pIIVKRkHnhjAUcBXz/s4rP7NZSDHiJ5TnsWt1Zw4bzB669twFiB4wDt
gekA8GEWZa1HgBzSYEwb+ADlH9tz6Jbf/X4VNx1MDLCuqhy8zcX+gxq//sVafngAABGzDxBNayFUIZzKgIFifr0h+NJXXdx0UKPVCkLgaosgsPiBW1x88ctl
/GVpA7UhgtgB4aOQEwEEVR3YMJ/l242aANMzqt1AUWJmcXWEADACY+/Zm/S6QeGBtrOQLtzVhdYm6rTon6RQ2OBGehaABxDOBnPDhz+bNlxXW2Rr8AJAF4Fv
J6Yn6P8HKQwfnqgoZx546fpOKRXO17caGsw8WMaC8ACgcsP3XMzoSKmkUCopMBfzEBGBUgTX7TmoiGWOfvLDA+0QyAGPdL5uxVdW/otnnnkBly6tQGuVO1KI
AGMs7rvvMI4cuQvGMKiPOxWFFwF0UXjKgGIWlEoKL730Rywt/RU7d+4IlR+AOiXWMp599re4447bsHv3BHw/+yYE8NSlfwSYpX88TRe2PKUyxWTHjlEAQKvl
wdoB4zgmRATfN9i1awLlstszlNK9Ul54QKDThQd1e4F0ubbjKBgjeOCBewAA77//AbR2uhToL9Su5J577sQNN4z2tP5W4UWiNiB3zAt6GB8AoHUJs7NHemfI
Ib4vPeO/fasKwgPtEMi/ktMvrkUErVax2I8rS0RQfSYYW4UHwkYwL3z8or0coZ/ig0hnIUTC817XEwioEHwqBLoLbwaPoOoM5QlaU1cbUVSIAGu7wTYzXkfP
MG+G/rooPGX1g6Gynufjgw9WCvUAWaKUwtTUBFxXd9/ULcADgO78mc/y0XHa9bRWWF7+EO+9dwmuqwuPBiNRiuB5BsyMAwem4fuMjC3NQvAQxLrBdnYMDJ8l
IsDISBWVigsiguOka8gjAWilojAyUutdZ4YOUcpmq8+6qOWz2qRgB1kwOTmOkZGhbQ2BatWFtRJfuQt1SbbEeeCBeAjE0ge3fLZJmAXlsrspWB7JCqW0R+SF
B5AeCfYv3LvCblHbvDxmux5tiOmBAvAiYQjE0vPB94/rlZVV2H5a5xDHURgfH0F244eB9M+CBxDvBvPt1bXbgFQ8lkoKFy/+GxcuFJkHdAsRwRiLvXs/g5mZ
XTAmuxcoAg+024DiG5WSNUdJutU2SL/hd7eeg8KH3eBWd2k7ElhLMDU1gWq1HIbAVpeMBY7jYHx8BMZ09wLoGot09N8MPuwGswtvBh9tVGYZmogwMTEW5BnU
Efosr4sEW2KZaal8eeABiY8EO4UHs3w2fJSHKFwUHdABBMHUt5f0mgxFJYrAd0aCBeF7qVsqAZdXBK3WgJsmAjgamLiR+nZ3PYsXhI+FQAH4DDBmoFQivPmG
we9O+VAKqTFGhoSu73vAt76jcfcRDd/Lsc2Wbg966J8FD0SLosgJj/iSYJKOCHj3POPDS4zRscEsqhygsQac/wfj7hwLScFkrJjlo9NYCGQX7un2AkgGvLWC
w0c0HAfY2AAUpW9RUqJ2UingzsMa3KubzxQJlqeFCsEDkuwFcsFnrAdQ+ETI2DjhgQeLzAUExuTcbC1o+ch8nf3lnJZvh0CGeY2JuWYOKbLTnJgDJBI2h2+H
QPv/jIv2dHvpyf8Jb5mHmyKZjWF/eADQhWK+3QpeCyKF4QWAKgTfvl5+N992SQxl88ED0nlcPjd8WCb6fNLSrrv9R074MH4VEu1VDsuLAAS8964PomAQxIzg
easeH97GDxHw7jtecfjwSzNjXSlVE7G53J4ZqFQIb77ewhfuKOPArdu7BLaZvP23Ft74cxOVikKw9DgofHhCCsJ2XQOyrJRzwPimswm3CXxUCRFgrOBXP/8Y
X/tGFZ876IarwCmJGs1kuHYdZUoq2VjBO3/38NorV2CNwNHRk6mDwwtEHNJkxC5riJx1lHuToRYDcAaFjyp0HIAt8Mrvm1B/aHamyZ0isWNJlI0SMhWNHgBL
XCNo8a0RVKqqEHworFRJAa2zGkIvs/CDItH0fTD4djVhPNaGw4UKCcydapzb/yd1EUhsVtW5SZSEQtRQBXkJBCsoBC8AhEEiTCB+WRPLGc+7sqydypQxLQYF
zw0NBB/LJ7aTmIRPkiTh44TJm98N38nHsUwF4FnrMnl+Y1mxf0bNL0ysAvR02R0mBnMR+HitaZcdBF56wYt0wUvsIC988MVcKg0TET29cOrmVarXRTWbK0MV
n97STmXG85tMpNS2wCN+3hu+b9uwjfAizFpXlfFbF5tu4/Y9e25pqHPnQE89NbnGkIcICkppEeH49a4XeCHlSPi08UPPP3/r2rlzIHX6NNkXj4lz4qeTr657
Hx+vuKMOQQkLMxIX/lTDM0iJWxp1/Nba8ZMv7H01epVWAUDwRqXoxxd2LzbXPzqudUWVdE0JsxFh+bTCC4uIsNG6orRTURsb/zm++Jv9i/W7lnT0HnF7LjA/
T6ZeX9InFnYtbviNe0Vwfqg2qR1dJmFhQKxIstO+FuHDZtMKC2tdpmplpxbB+ZZ35d7FF/Yu1u9a0vOv3W2iy3W/PB2+SDn3vbdHxyZunCPgYUXOPpCCtT6Y
fYjwNQkPUlCkoVQJIgxm/wILnrtsLi6cOvWV1aw3yDNn9fGMj/7g8lipoo6KyFERe0gg08I8hMS6aMz18sIjdQPzwicKUANQywQ6K6TObDQaZ355en/f1+f/
Bx79ZhrAgFJCAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAVhUlEQVR4nO2de4xc1X3Hv79zzr3z2Je9NnFMDF4Irzi2
3OAKVGhrmUitFB5S1M4mUpUSaP9LoamiRE3zGBZCqHjkYYciRSoEQpJqR21SE6rmIRBFJAKJJATiiBBiTBJwsA37nsc95/z6x527O497d2fGM/fOrucr2Ttz
zz2POZ/z+p1z7rmE09B0juUvdoGnpsgG1750a+lCgC43jMuYeTdgd7LlrQCyDKi6AJhXPjYGzo0fOfRGDv3CYV+b/XGIw7JbeIIivDS4ccT1Zg9hTgRoMC0x
0UmAjpGgF4jpGZB9+rN3jb0U3JfPszhyBFQokIlO1eqiTjzl8ywAIAD/xVtLFxOL9xu21zDxXkemh0kAbAFjNKzVYGY0/VxuALV8vfFjD+DXxN/kFgd8ABzl
yAQiAgkJIRwIErBsYHR5gZmeIyG/S2S+/dm7xl4Emnm0ozYLANP0NMTkpF/ivnBb8SoBeZO19up0OuVoDVS8MthqywQmC2KAQNQc1wB+hMPKL/bvYSYCswUJ
IYVSWSjlolSe9Yjko5b1oal7Nj8GANPTLCcnYQFaJeZ6tVwA8nkWQQm7+9bKPkk8JUhdLaVAqVQCs9EABAgEEEVwC35auNsAfmMoDX+Yq72GJRIq5Y7AGg/G
6kcN2fxt92x6FqhntZZaKgD5/ONqauqAzl9/ND12wdm3CMLHlHRVqbhkLcBELFCt5iHpbviRA/jhDmvBb7qfmdkCoFRqVGhd0gy+55VTP7nlwQcPlAJm0Snx
tWYByOdZTU2Rvv2Tp3ZnsyNfS6ecfYuLRT9yIhlFaAA/zE+X4Ne5MZitIRIimxmncmnh2bKZ+fAdB899IWAXnaI1CkAQwN35xeuU43xdCme0VFrUREKFohzA
X8VPb+DXXrSWddodVdZ6c54pfehzXz7r8FqFILIABB7v+uzCDW4qfb8xBsZ4hiDkAH6z52Th11xna4R0pRQOKt7Cjbcf3PbAaoVAhF2shZ/ODN2vtWet1nYA
PzzCfoEPACAhjfGsp8s27Y7d/88fef2GqSnS+TzXz8EEtzde8E0JMnd8ZvbaofTwYU97xlotCIIG8Js99xP8ug7BGhZCWaXSsrz01nWfv2/HI9M5lpMNk0Z1
LUA+z2Jykszd+dIlaTf9TW08ZmNoAD88wn6FDwBEgqw1ZEyFVXrkm5/8hxOXTBbIBJNGy/etBMBUyEG8/qdQ3lvFp1NuZm+pvDjo89ch/NrEWbYm5Y7Iirf0
3Bxvufz4cehCYWWyaLk0FKYhJgtkSqcWbh3KZvaWSot6AD88wvUCnwEQCVkqz+tsenzvkP3DrYUCmencCncCgFxuWhYKOXtnvrLHUfiJNRbMVjCIBvDXL/ya
L0wkrBAKnjWX3nFw/PlcbloUCpNGAMCuXTkGiIkrdznKldYaDOA3e16n8AGArNVQMi1hyncBYJ85QMGo/wtTS3+mlPt/lUrFMDCY4ds48Gu+snFURla84p/f
8ZW3PTmdYykKBd/RGP6EFLK6BDWAH349zM+6gQ9msBAOYPkTAFBAdQxw56dm3qnczC+ttYphgdr5gQH8VfysH/grHwhCCM2s33X7oW0vCwAg6U6m067DbAwG
8EOuh/lZd/ABgJitSbtjDltMAoEZSHyt1gx/LT803aGRD+A3R9fH8AEARETaVMBE1wIAfSFfPF/DviBIZpg198NmDkLd7oKGW9uEH+rWGFQP4Vc/Wg4JN2b4
QTKJBLG1RbZ2tzJsrky52UylsmRBJJKEH0D3KoD2VrKFmz40BNNF+M0gVxwiC8YaNR8EKAU4Lq3EnQD8IDXWGus6Q5lKZfFKxRCXCUFgAtNq/noMXwgfvNGM
bWdL7NgpMTJKKyMS5qZmgRG2+60BfsiCN3PjQIfW9IPqXqwIpwYHBjiAzZibZfzumMZrv9UQAnBd8luEmOHXlEcWJEBElymC3WMsQNYfASQFv1RkjG8VeO/7
0rjwXQqO09GG5b6V9hgvHvHwvf9ewonjBqkMwLW79mKCD78rIn+ntt2jGNhpjQaHF+5Y4JeLjB3nSnzghgyGRwS0ZlQqq3TE61BCAO/e62LinQpf/+o8Xvm1
h3SGYC1ihQ8AICZrNQDsFARssVYD1DjsQs/hEwHaA4ZHBf76b3345bLvLsTG+gcA5TJjaFjgb/5+BJvGBTyvvtbFAh8MApFlAwBbBDOyoc+nxDTar5QZVx5w
MbbJhy9lc1I2iqT0C8HIqMCBv8yiUublYU1c8ANZf0yTFSCSTaZNHHY+AGOAkVGBS/YoWMvLNWUjSwjAWsauvS42bRbQGs2db4/hczUOAsnmLI9pkofgD4zG
zyKMjPp9YUgntOFE5Bf84RGBs7ZJeF7D6CsW+Cv31heAOGf4yI8ulSYIgWgbe4OKCEhlqN4kjRk+UFsAkp7ePZPFSAQ+AKjIiBsu9Ap+648xblRx9Lcewwdq
WoDkaj6jw6fUN4aa8jnsS2/gM6oHNiRW88+0jr9RCTX7ftS+g0i25g8EJFPzA4kol57CP9Nrfo0SgV+T/6LepfZjPPAHRaGqBOADQQEYwE9WCcEHAJV4s99m
KbDWX0MNW7uKS8sDqE7nrjniM+KFD9Qc2xY3/NpdMa2KmeG6/bNa5Hm2qwUxbvjMdWZgjPDRvuXPzHAcgZ/97Jf44Q+fwvz8IoSg6D16PRARwVrG8HAWV111
Bfbte3fXCkES8AFAJQG/3RJgrYXrSvzgB0/hi1/8dwCAlCJW+IH8QmDx2GM/wk03fRjve99+VCoWQnReCJKCDwAqEfhtiJmhlMTMzAIeeug/kU67SKVcWJvc
EFIIQrns4eGH/wuXX/5HGB8fg9adtQTLOw0TgA8EY4DE4K8NkdnfJ3Ds2GuYm1tAOp2C1h2fjNoVWQs4jsLCwhKOHv0dtmwZAzO3XQBay8vmm7oFH2CIJOG3
Ng70M3V4OHtazWy3ReR3B8PDQ6cXUILwgYbDm2Ov+S2UACEInmcxMXE29uy5BD/+8U+wZcumRCcTiYA335zFvn17cMEF50Br25lJmDB8cJMZGB55L+C3C5CI
cPPNN8Bai+eff3F5PiAJCUG49NLd+OhH/w5CSBhju7KbKU74y1ZAU7gxwW8nw4gIxjC2bt2E2277GH71q1cwN7eQyGQQM2NkZAgXXXSe/zBLF8zAuoYwRvjA
shkYHnk/1PxARLSc2RdfPNFZIF2U1gxj2h/4NWq1/F/+1CP4QO0YoI/hBwoyu1JpeJoiAQkhutsCJQAfaDID+xd+rXxroH8sgk6VVLNfK5Ek/BaMgI2vBOED
y/sBEqj53N1WYV0qYfjMgEiq2WfUPZR9Bis5+ECjGRgXfA7MwM7gc/MD+X2gTiyClcqXBHyAa8zAGOFHRriGmP0BoJT9Bh8A/MfbrOW25jjWyo5ewgciFoPq
E9Yf8AFAKUKp5KFYLHfkv9fKZFJIpx0Y0/rvS6rmB1KJw28hr5h9+CdOzOA3v3kNxiS7GhglKSUmJrZj27bN0LrFliBB+EDjmzwRL3zm1toCIfya78O3kFIm
shlkNfnT1RZHj76G0dEhpFJux2mMCz644bmAuGt+K0M5v98HisUyjDGJ7QRaS8wMKQWMsVhaKlWfeO5gnOMHFhFH6N2nlf/Nq4Ex1XxQa+N4In8DRibjQkpZ
bQFEWIISlr9VTEqBTCZdPe+ggw0iMcIHmszAmOCHJytS1jLSaRcTE9tx9Ohrie8IChORvz4wMbEd2azb8hiAgBozMF74QO1qYMzw25F/qgZj27bNGB0d6ksr
gNm3ArJZt7pK2KK/4L8E4AMN5wOEJqxLkXej29aakUq5yGTc0w+sB7IWrY/+EXCPzphewwdCrIBlr3HA5/YGSkT+/Vr3W//vi4ja2xnUUl7W39xN+ECUGRgD
fAqPqCUl+VhYd5VMs1/rodkMjKnZZ4SeTnjGK66aH9xdf0RMXPC5c/j9OAcQ6HTTFhv8mvxfMQNjhN+piAhKVc/X7UMJQTCm/d8YbgT0Hj4QmIHrAj5gjMGr
r57A4mKp7w6VZAaGhtI4++yzoJRo3ephdDX/24EPhC4GxQy/jTLx0ku/xcmTM1BKdsWs7KaIgFOnZrC4WMQll0y05CeUfYzwgSYrIH74a3EMVgLn5pYwMzOP
VMptirtfJKX/EOv8/BI2bRpae04gwZof+BF13zqMvBP47R4Q2W81fnV1ktj44QPLLUACNT/q7SQNCqaBR0YyGBsbxqlTs33bBWhtMT4+gpGRbFvTwUnBZ9Sc
D9BPzX6ULrzwHGSzaSwuFtGPewKz2TR27Hib/+6ttkaBlAh8IDADE4MfPgyKSodSCuefv72PzUB0ZgYmBB9gqGThtydmhued/vN4vVJHzwomCB9oPCm0hci7
Cr+DvqBf4QMdpG0VkHHAb7AC1o48afgbTUnDB8JWA2OCH/mypDXEvHpauiH/+JfexhGlOOEDjUfExFjzWx/+1adDqXjgeF78hSBu+IzaTaExw29XzIDjADNv
Md54vb2nb9qNZ2SMsP0d5L/RKyYlAR9oMgP7G/5Pnzb4zrc8lIodBNKGSABXHJC45q8c9Pw4Qk4OPoBaMzAB+C30A8z+Cxdn3mR85z88lEpAKt3DcUB1m+5j
/6Nx3gUCey5VqFR6+E7DBOEDDecD1IQfqa7Cb0HM/gTLG8cZpSUgnfY3XwaDwa7/s358ygF++0qvq3+QDcnAB9ceFbsSfnRiu9jst9qFEwGWgbO2EVJpoFwG
XLeazm6PA3jlQRTtAe84t9ejwKAJpLpLccEHIq2AsLi72OdXf3MrG8OIAKOBzVsI133A8buBHo8BQMD+v1DY/R4JrWN8pW3M8AEOswLC4u4yfLQGPxCRb5b9
8RUS510k8MZr3JO1IILf2oyOATt2CsT6EHIC8BktLQYlCz9QUAjGtxK2bO29gR7rPEBC8AGstRjUO/jUcL0VEfl980abCUwKPhDxZJAfd29rPkfevLqSnKbt
uoIxYKhb7+GDIxeDYmj2e1yL14Mi+ccEHwgpAHHBp/pbBgoUI3yg8dGwmGv+AH6DYoYP1D4aNoCfnBiJwGdUzweIHoXGAf9MLwpc83+oU8/gA2FvD6+5o5fw
lw8KPZP5M4fkc+BW9yfswmnDB6LMwDhqfrXZO9OVGPxq3jebgbH1+QP4Sdb8QKrxjljgV8cdJIBSif3zdcPC3qiqPhVVKjJI1PGr/RN2oTvwa1peUXtHXPCD
P1IRTr5hMD9nIeSZ0SMEG1wW5gzeeF3DcQioeQNOnPABQDCzAVOs8AM3pYDZGYtf/KwCIfr34Iduylr/IInnf1rCm6c0lFrJlnjhExhsBJiW6h9oiAd+cMl1
CY//7xLmZi1SKYp3CTZmGQO4KcLCvMX3D8/DdQlcLfRx13wQgRlLgsGnSEgwM8cJHwDY+ps9Z9+yePirc1hc9AsBgOrZ+xvnHzOQShFKSxYP3PsmTr5h4DiE
mlwPzfhewLdgJpIA45QC4ZgQagIos18sIiKuS9/pww++WAukM4Tf/LqCf7trBle/fwgX73bhuhtrWGgM4xfPlXB4eha/f1Ujk/G7vLjhc9W7IEWGKscUIJ4X
JPcTrYQaF/xAxjIyGcLJP2g8cO8sdkwo7Dzfwegm2bD0G5Ky1i6t4cir+1tjcBoe5Mpgd3bG4pVfl3HsqAciJAsfvhEiSIJYPK9g9TOWDbj6BtS44QdhWgN/
ROwAv39V49jLOiitDR5XUsFcHx4BYCJ/XSPCpAi/zKu4+c6hRWTVClOfW45DfvfGSBQ+/KwhZgPAPKMIzlNeebEohMxYtozG3XYxwF/+U/3rugQ3FRU41f2Y
pqiX/2vuQqLhr3JAQx18inBbTpp/gRmNMxvBlvP6ZMcPH2AWQoqKt1RUip4iAMj/08kfue7In5QrCwaAbE5f7+FHJ7j+pvpWIcRfyzV/jVpf59b8+6O7i2RW
9WrjjwgikHGcIel5Cz8+9NDOK0TV7REp3aol0Ji+AfxGr+sYPpiZpXBhYR8BqjOBRprpUnnWIyKJOj4D+I1e1zN8AExEslyZ9QTENACIXI7l7Xdve9kY/b10
apTYwoREPYC//uGDwcZ1R0gb73uHHtr5ci7HUuSWncWdRmsQMQ3gbzz41QtkjAfB4s7gqpgskMnnWdz2pfEnK5X576dSo5LZGoQENIDf7Gm9wLewxnVGpOct
fv8r39j5ZD7PolAgIwDgyJECAUwC6uOeLhkhVFNKBvCbPa0X+AywgILRJSOF+jgA8plXB4GFwqSZnoaYOjj+80pl6Z5MerO0dqUVGMBv9rSO4IOtNWl3k/R0
8Z6DD53z81yORaEwaYD655Ipl4N4+9uhNolTT7vO8N5yZc6AhKwPcwB/XcFna1w1IrVefA6bypcfP36hLhRgg9Oaa7aEEe/aBT50iMra8AeNrSwI6QrmYP6q
PvIB/Kbo+g++ZSvJFdZWFljYDx46dFF51y5w7VHdTfOl09MsJyfJ/MtNr1+bcscOa1My1mpR+z6sAfym6PoPPlsWpKyUKVkuzVx337cueiSXY1koUN2Oi6ZN
oZOTZPJ5Vp8/tP2Rojd7o6uykiA4aAkG8Jui6z/4li1BsaOysuzN3Xjfty56JL+fVSN8IOLh0Kkp0vk8q389uP2BUnH2RkelhZSOsBHmYd3PGcBvvBJyobd9
vhSOUColypW5G+97+IIH8nlWU09Q6KF3q+66yOdZTU2R/vRHjl8n3OzXpVCjpcq8FhD1R8sM4DcFHT98BlvWrjOiLOs57S196N5vnH84vz8aPrDKYdHASkvw
uXvffnixOHOlMd6zw5mtCgCzv6A8gN8H8C3YgIkz6S3KWu/Zsl640of/+KrwgRZP2vEDOqCvv/7x9Lmb33MLmD7mqJQqlmctCExMojGsAfzGC91f2GFYywxK
uaNC67IG+J7ZV1685cEnDpT2739cPfHEgTXPOm15410+z2JqiiwAfObmE/uYnCkp5NVCOCiX58FsNQABwsoZHgP4yx+6A5+Dd25aIlKOMwxrNazRj4K8/MEH
dz4L1LNaS+0ecE/TOX/9AAA+9dG3rhKWbmKYq1PumGNMGRVdhLXGVt/zTCAm8qNZ9XXBA/ho/EFs/cRykJdCSKFUBlI4KFfmPJB81Gp96CsPn/MYAPhm3sok
TyvqaOttPs8CAIJS9ul/PHExsfN+C3sNM+9VIjVMQoCt9ksoG1jmVQBHpzcW+FUPicFv+jF+IypIgkhBkASzgTblBRCeA8nvEplvf/n+HS8CzTza0Wntvc7l
WO7aBa6N+NM3z14I4HJAX2Yt7wZ4p2W7FYwsIt9W3l7NX60Wb4CaDwY0A0vE4iQRHSOiF5jEM2Dv6YNfO+el4L58nsWRI6Aw+75V/T9sU8CY+BUofAAAAABJ
RU5ErkJggolQTkcNChoKAAAADUlIRFIAAAEAAAABAAgGAAAAXHKoZgAAKr9JREFUeJztnXmUJMV957+/iMyso3t6ZjhmhkMwMBzDDDJgAc9CskDy82L0LNtY
7l5bh/129YQPMF6v/Z4NQqrplYSEbIyQhpGEFlugA++0ZMmS1jJ+skGSLcsg9qE1DGMNMCyIawaYq7uOzIz47R9ZWVPddWXVVB+V+fu+10dlxC9+UZn5+WVk
ZEQGYQWqVGIFQAGw09Nkm9N2fITXVqvBJka4GYo2M9tNDJwKYB0s1jK4SIwCiBQnccatubradUnkTpkS2fRK4O7JPfx0LjbRXjpql2ynzv8vmYv5JfT4/sls
kjlKeJ5YgCoAykR0gIF9YPyESD1pCbtBzu68Dp684aNrDjSbdTuXV4JouSsQi5lp2zbobdtgiKhxTG4pzZ2stb6ULL+BwRcDdA5gN7hOUSkVHXTLDDYGli04
/knmtHVT1/xJkgT+ef+lAf66CApE8Y8GKQ0CwbJBGFQsQC8C9scg54eK8S9G6wen/2zs+aPfgmlb6QG9bdsV887x5dSyB4BSidXWraCpKTLxto+XauezpqsY
fJW15pJcrjCuFGAMEIYhjAkAtpYJDAaIQQwQEQAk3LMCf28jCPyt1WC20f9MVHfNICJSWrtwdA5au7A2RK02O6tIPQTobxnwt6ZvXfVoXMrOnawfewy83K2C
ZQsAk5M79ZYtk40d8Bcfmj1JGe9qJv51a/myfD6njQGCoAZjQ0MAM4HAUFRHHUA/zDVlEviTSOBvrUZnOwYzM1HUIAVAihztugVo5aJSO2yI6PsE+mtt1Fdv
+vj4C0B0Ady1a4ZmZqZMx6IXUUseAOr3RIjB//j03CUg970MfnvO844zBqj6FYBtyAQiQAHUvp4Cf08J/K1GiwF/+4/cuEMlIsdzx6C1i5o/9yqBvhLCfPaD
t655CGjlYqm0ZAGAmWlmBipu6t86Xb5cKeePmPkXc55HtZoPYwJTZ131rJvA31MCf6vR0sHfrgi2zAytXO3lVsH3ZxmkvmlMcOsHbzvuOwCwc5L15E7Ypeoj
WJIAUCrd70xPvzkEgFtL/qXKwU1E6m2O1qhWywAjZIKOrvT9n9kCf5JiBf5lhL+lMswwRHBy7gSMDWCt/Qaz+dD0bWseBOYzs5ha1ADQ3Kz50I1zp4wXvRLY
vsdxPFWplhkMS0QaiHeWwN/TTuAfyNHKgX9+JmY2AKt8fg2FQcWC6K6arU3ffNsJzy3FbcGiBYDmCPYX05XfUUpPe567rlwpMzMsIQIfEPjbbRD4k9gkc7RS
4Z/3ka0BQRXyaynwy/ustaXpj6/5NLC4rYFFCABMO3dG9/o3lw6eVXSK2z3PvbJaC2BtEALktO5kgb+nncA/kKNRgB84Wk9mDrV2nZy3Cn4wd1+1UrnuozvW
P7FzJ+upKVhguH0DapiFxU2WqSkyt5bK7yo6xQcd172yXCmHxgQs8LdLEPj7t0nmaNTgBwAicowJuFx5NXSc/JX5fPHBm67f/6648zxmbFgaWgsgilBkrrnm
h+55p772NtfzrvX9qGefSEXNfYG/6waBP4lNMkejCH9LKlujladdt4ggqNzxwtyeP7zzzouDyUnWMzM0lHEDQwkApRI709MUfujGl08ZK6y6N5/zfnZ2rmyA
pkE7An/XDQJ/EptkjtIAf2zDbBlEdqxwvK7WjnzvcPnQb9z2mdOei5lL4rqbjjkAxBX5WGn2Qs/xvuq47sZKeS4kpZxGJoG/6waBP4lNMkdpgr9Zlm1YyK9x
wrD2tG/KV998+/pHhhEEjikAxBX4s/cdeYubz/0NKWe1H1RCAgn8HRME/v5tkjlKK/yxLHPouUWH2R6q+LO/esv2k/7pWIPAwAEgdvzR989emfO8vwWQC8PA
xM/1AQj8An/HMpPbJHOUdvjjzczWOE5eE6jmm/Iv33z7+vuOJQgM1KPYDH8+532DmXNh4FuBv1uCwN+/TTJHWYEfAIiUDsKqtWxyni5+48Zrn7tyeprCUomd
9tbd1XcAiOG/5f1H3lzI575mrXWMDS0pdbQsgb/rBoE/iU0yR1mCP5YipUzoW2sDJ5eb+NoNv/vCm6MgcH/fQaCvW4Cdk6ynZsh85KbZC/J577tgTISBL/B3
TRD4+7dJ5iiL8B9NZDBb6+icIlKHK8GhN92y/TU/6vcRYeIWQKnEamqGzMdKsxtynvt1RXoiDAMj8HdLEPj7t0nmKOvwAwCRUqGpGSI1kXNWff2Pf++lDTMz
ZPoZLJQsIzNt3Qr6zDXsatJf8TzvND+ohnLP3y1B4O/fJpkjgf+oiJT2g3LouoXT8q7+yjXXsLt1FwjgRK37RAGgtA16aorMwZPmbi8W85dVynPyqK9rgsDf
v00yRwJ/q4jIKVcPhoX8cZed4L10+9QMmVIJum3mhba9MsRDfG/5wJHfGC+Of6lSrYQABP6OCQJ//zbJHAn83W2YERbyq525yqvv+Oj2DffGfXbdiu3aAiiV
WE1Nwd5Wqmx0Xe/TfuBbgOWev2OCwN+/TTJHAn8SG1Z+MGddJ//pP/ntFzZOzcD26g/omrh1KwggDtje5bneRBgGDJBa6F3gb90g8CexSeZI4E9mQ0TKGJ9d
XZjQjroLII76AzqrYwCIm/4f+8Dse8fHim+pVMuhzOrrlCDw92+TzJHA34cNAwSlq7VDYbF4/FtuuG7fe6dmyOyc5I79AW2jQ9xsWOvNnmiNu4tIrTEmAEBK
4O++QeBPYpPMkcDfh03z92dYrV0w24MG/hZv7af2A9vavlqsbQtg61bQ9DRZv6Zuzudyx4WhzwJ/uwSBv3+bZI4E/j5sFiQQQRlT43xu4jhYunl6etp2uhVo
2VgfSWT//H/MXugo94ehsQxmBT6aV+Bv3SDwJ7FJ5kjg78Omy5cmIkvkkm9rF9/yyXWPTE7uVAsXIGlpAWzZAgbAHNLNjuMptgYCf/cNAn8Sm2SOBP4+bLpX
gqw1cNycUmxvBsBbtky2WMxrAcQdf7dN+69XDn0/CAILPhokBP7WDQJ/EptkjgT+Pmx6VCJmlZmt6xYU2+CyD95+/L8unCswrwXw2GORXWCD92vlgK3lhQX2
lMDfUwJ/q5HA34dNQvjjj1p5CMPg/UCjhd9QowUQR4bo1V7uw6GxAEdPAwT+1g0CfxKbZI4E/j5s+oM/dmEd5cKE/utu3rH+kcnJnTruC2i0ACYn4//oOs/1
FKy1nQpMUjGBP0mxAr/A34fNAPCDATBb1xtTlnAdAEwehT1qATAzERHf/mE+MfSre0ipCWNDyFp9rRsE/iQ2yRwJ/H3YDAp/3ZmKbukPW8qd/ZFPTuxnMBEo
auJv2xbNHPL98mShkF9tTWgE/tYNAn8Sm2SOBP4+bI4JfgAgMiY0+fya1cSVSQDYVp8tGN8C2LrRO41hMCDwC/wD2CRzJPD3YXPM8EciAhnjg5neWd9kAYBK
JVbT02T/vFTdTIRHLbOK0xJ4TlrPeiaBP4kE/tZqCPydkhN/YQYAImUt8/kf2X7i7lKJlUKjFWB/KZ/PabAxEPi72wn8AzkS+PuwGS78AEDMbHK51Rqwv1Tf
phTql3sGv9VY1Jv/PT0nrWc9k8CfRAJ/azUE/k7JA3xhAhkbgBlvrW+xBAAfvuHIetdVTymli8whRx2AvSsm8CcpVuAX+PuwWST463ZM5JBlU6aKPfMjd214
SQGAm1c/UygUi9bGvf+9KybwJylW4Bf4+7BZVPgBgIg5NHlvokg5/TNA/d1+ZPBGrbp0/Qv8bdUcKhlYMGOKe75xMS6WWgrobtGwS2TTZM11N/Nsep91XK9k
Ylf178/92lDLv+0KbrGJd7fA37tYtmBFDljxGwH8bfRyT8IlxkSnbMuOF/gBRDsm3jnWAmzjZ6cL7Pq4VK74K39SqNo4WeorP6l6j1a9Cm0PQ8bhrwNO1gZg
5ksAgD76J7xa58q7teNtCMMaEzXFAIG/AX4QAIHPIAK8HMHLEbROMfz1LCsjALTZ2rTJWKBWZfg1hrWA6wGeR/MDgcBf/8PsqBwZW33RBt5mh/Lls4mwzpgA
Av/8jUoBvg+YkHH8OoVN57g47UyN409UKI5FAUC0fIpvFUzIKM8xXt5v8fSTAfY8HmDfCwZKA7kcwba8GDub8AMAgchYHwCt0271bIeM3ezmx5Tvly2o9Y2/
CX2mCn6q/yrPMTacovH6yz1sPt9BvkANq47NTNGSi4iwei3hpFOB117koVZlPP7vPr73j1U8+3SIQoGige315lNW4Y83MLP13DEV+HObHdJqs1IEpvpaQgI/
rAXCgPGmn8/hZ3/Og5cjGMPw/ciiw3MS0bKKG6eh4wIXXpLD+Rd6+O63q/jHvytDKYLWDNvrhE45/ABABFZKA0SbHWbexFzvUe6vnvVM6YLfmKjp//Z3F7D1
AhdhGIGvmjqZRCtTcWBmRuOYveWqAk5+jcb/+twsfJ/hONS55ZYB+CMRmC0Y2KQIeA0zAEY/D4fqmdIDP+pNembg134zgt+vRTkF/NFTfMxqNcbm8z385u+s
gtYEazu04DIDP8DMFAUA+xoFYJ01BrzwMXZPr+mCX6moJ/kXfjmHszc7qNUYSjr5Rl5aR0HgjLNcXP2O8UZQn6cMwV8XGTYAaJ0CaI1lc7T51NtnKuGvlBlb
L3Rx8WUefJ+lhz9F0jq6JbjokhwueUMelTIfbdVlD/5o8BRbALRGMbhY/9CuG6CN13TBD0Sdfrk84YorPTCzdPKlUESAtYyfe2sB4xMKJkQm4Y+7AZktiFFU
xCjUOwQSeE0f/EoB1QpjywUOTlyvEQTSy59GEQFhCKw9TuOiS3KoVrht30764a//by0YXFAgUlmFH4i+ltbA+Re5nZ2LUqH4JXcXXOw1Rgo2Kyvwxx+ISGUa
/sZV4XiFU16jYUz7q4IoHSICbMg46VQHJ27QjaHdQPbgj9X7dE8p/AAAioaRnrBeIV+IHhGJ0i1jAdclbDhFIwzR++2XKYYf6BUA0gw/jo76W3NcfQR0ouaQ
KA067gSdYFRguuEHugWAlMPfSGKgWJRev6ypONbrZQ3phx/oFAAyAn9Dwn/21O2YZwR+oF0AyBr8IlGzMgQ/sDAAZAx+TpZNlBVlDH6gOQBkEH6RqKEMwg/E
ASDr8Ev3f7aVUfgZgCPws4z9zay44zmTBfiB+mvBE5eRRvhF2VSXDqBMwF8/91VLQlerXkkCv2i0lQn4m6Q6JSSow4IkgV802soM/E3nv+pZhsAvyoCyCD8A
dJ8NKPCLMqCswg90nQuQoECBXzTiyjL8QMe5AAkKFPhFI66sww+0nQuQoECBX5RWZQh+oGUuQIICUwi/hAgRgMzBD8ybC5CgQIFflFZlEH6gMRcgQYEphj/Z
+9BFqVVG4QcAlXX4RaJuSjP8QJfHgJmBXyJBdsVIcD6nF37mNpOB5vtOMfw9Dv6wZa3N/AMJIkCNyHvXswA/0HU2YMrhXyJZy1CK4Hmy2CAAhCHDWruiA0FW
4AcWBACBf7iy1sLzNJiB3bv34skn/x+OHJkFZez9A8yMsbEizjzzdGzevAmep+H7ZkUGgSzBDzQFAIF/uIrhf/TRPbjnnq/gxz9+CrVasLSVWFEieJ6LTZtO
x7vedTV++qe3rLggkDX4gXoAEPiHK2sZnqfxwAP/httuuwvWWhQKeeTzuaWrxAoUM+OJJ57GBz5wK6699rdw1VVvgu9bKLX8LaIswg8AjsA/XEXwK+za9SRu
u+0uuK4Dx9EwxsKYrF79j6pQyMNai+3b78a6dcfjda/buuwtgazCDzQeA2Yd/uGBqRRgjMXdd38Z1toG/KJIcQeg4yjcffeX4fuhwJ9AiwE/ACiBf3jwW8tw
HIUnnngG//EfT6FQyAv8bWStRT6fx969z2LXrifgOAS7DCuzMhhZhh8YaDJQr4Rswg9E97gA8MQTT6NWq62Ie9uVqmhp9hB79uwF0Pe5vTjKGPwA9zsZqFfC
aMI/7HPvyJHZIZeYXh0+vFz7KttX/rj0PiYD9UoQ+GONjRUXodR0anx8BeyrjMIPJJ4M1CthhOFvcxs4qOLxPZs2bYTneViG29qRETOgtcZZZ20EsIxrs2QY
fiDRZKBeCSMOPzC05cGVUghDxjnnnI6zzjod1WoFWq+cgS4rRUop1Go1nHbaKdi69RwYw8vzJCDj8DN3CACZgn/Iih/9vfvdvwprGcas7HHvSy2lFJgZtZqP
d7/7auTz7vI8KRH4AfSzNJjAn0hKKfi+xYUXnofrrvstbN9+N7RWyOfzmV+CkBmo1Wqo1QJcc8078PrXX1QfCbjyAmQW4Ac6Tgbqat9aYCKbJMWONvyxlCL4
vsEv/MKbcMIJa3HPPX+DvXufRRiGi+98BUtrjdNOOwXvfOev4A1v+OllHwHYSVmBH2g7GairfWuBiWySFJsO+GNFLQGDiy9+LS64YAseffTH2LNnL2Zn5zC0
TocRETNjfLyITZtOx2tfuxm5nCNX/oVZlwF+oGUyUFf71gIT2SQpNl3wx4qDgNYaF110Hi666Lylr8QKkzGQK//CrMsEPzBvMlBX+9YCE9kkKTad8MeKO718
X94IFL8RSOBfGfADHV4JJvAPXyvxpBdFyir8QKKVgQR+UXqVZfiBnpOBBH5RepV1+IGuKwMJ/KL0SuCP1GEyULbgz3rnXOaVUfiBtpOBsgd/tp7Ki+Ypw/AD
LZ2A2YNflGFlHH7wvAAg8IuyrOzBz+jwUlCBX5QtZRN+oM1LQbMHPw9WkCjVygL8QJLZgKmGP5Z0A2ZT2b3yxxtUU1ovnwmqhRGCX678otbjnyX4gXofQDbh
F4nmK2vwA4AS+EWibMIPtJsNKPAPXSyPHQbSUi2jnlX4gYUBQOAfqpijOfCOozL/PsB+xQwYw419uGh+On5IaDfC8APNAUDgH6qi994TrGW8+uoRlMtVWMsS
CHooAp5QLOawevU4HCd61fpi7Lesww/EASDj8A+7hc4MOA7h0KE57N37PObmqnIb0KfiILBx48lYu3Z80YIAgMzCDwBO1uEHY6jDAJgZWiscOlTG448/3Vgn
QNS/KhUfjz/+NM4773SsXbsKYWiH2y/AyDT8QKdXgnUrIkXwc8sOO3YRRUtd7937HKy10FrL1X9Aaa1gjMXevc9jfPysZX2tWhrhB3dZGqxtESmDf9iK7/sP
HZrD3FxV4D9Gxa2pSqWGgwdnoTUty/5MK/xAP2sDCvwJFN2nlstyzz9MMQNzc5Vl6UBNM/yMpGsDCvyiDCrt8ANtAoDAf2xiBgqF3JINYsmCiIBiMb+kPrMA
P7AgAAj8xyYigjGM1avHUSzmYIyRQHAMijtUczkPa9asgjG8JPszK/AD3dYGFPgHEjPDcRQ2bjwZjz/+NIyx0FoWBRlE1lpYy9i48WR4nh7+Y8A2yhL8QKe1
AQX+gUVECEPG2rXjOO+80/HUU8+jWq1J30OfIgJyOQ8bN56ME06YqA8EEviHCT/Qbm1Agf+YRYR6EFiFn/qps3Dw4BEZDdiHolGAeaxZs6p+5V/8IdRZhB9o
mQwk8A9LcRDQWuHEE9dg3brlrc8oyhgI/PUiFgN+YN5kIIF/2CKK6hKGdrmrMpIiokWfCZhl+IHGZKAsw8+LHjHkScDKU3TE68clo/ADgJNt+BMVK0qrBpgM
kib4gQSTgeYVmDr4eX4zUJQdcXzkk7fOUgV/vWI9H1CnHX5AXgou6q1Uwd+kZLMBUwy/SNRLqYO/qXK9ZwMK/KIMK83wM3rNBhT4RRlW2uEHus0GFPhFGVYW
4Ac6zQYU+EUZVlbgB9rNBhT4hy6ZAzCYlmMAVZbgBxbOBhT4hypmWRhkUDEvzcIgC30myDWATWsRKwF+oHk2oMA/VDHLwiCDKAJ+aRYGWeg3Qa4BbFqLWCnw
A30sDCLw9+VeFgY5Ri3lwiBZhR9IuDBIquHnRFXpw70sDDIsLfbCIMkPfTrhBxLNBUgx/LGGeGWRhUGGp8VeGIQTRYD0wg/uORcgxfA3Hfxh4Rnf98vCIMPR
oi8MknH4ga4BIOXwt/47BMnCIIsh5uVYGCT98AMdA4DAL8qysgE/0DYACPzHImZZGGTYWtqFQbIDP9ASAAT+Y5EsDDJcLf3CINmCnzHvKYDAPwzJwiDD09Iu
DJI9+IFGAMg6/MMLC7IwyHC0tAuDZBN+AHAE/uGTKQuDHJuWdmGQ7MIPLJwNmFn4h392ycIgx67FXxiE0XzsswY/wE2zATML/+JJFgY5Ni32wiDNyiL8QDwb
MOvwL3IskCcBK1tZhR8AlMDfV41EKdLAcTkl8AMJ1gVoMRf4RSkR8wDnZorgB/cZAAR+UVok8Ed/EgeANMMvsUDUUymEH0gYAAR+UaaVUviBJGsDZgB+6aMX
dVSK4Qd6rQ2YAfhFoo5KOfxAt7UBBX5RlpUB+Bkd3gko8A9XNiUDAYmW7h39y6qMwA+0CQAC//DEDJACPA9IR08DI/Cj75RaZQh+YEEAEPiHJ7aA6wFBADy+
y+LF5yxMiJGMA8zA2DjhzLMVNpxCMCZq1aSuNZAx+IHm2YAC/9DEHMG/d4/F1+4N8PyzDGtHv+chXyC87vUav/hrLrSTsiCQQfiBeG3ArMM/RDZt/cq/9wmL
/3m7j8AHCkVgJC/9C8QMfPcfQhw+xHjXNZ7Anyxnk83Kgh8AHIE/cY0SSSkg8IGv3Rs04DdmuD6WU2uOIzzyoMHZm0O84S0ufJ8x5PU6llYZhh/o8lLQTkol
/EO6klkLOA7w1H9YPP8MI19IF/xA9H1yecLDP7AwhqFGuRWQcfjB8wJARuEfugjP/8Sm4p6/nZijIHfwVYsjhxjaGfC8WHHKHvxAIwBkHf5UnMFLqsZsuhFu
AfSFZgrhBwAl8A97KXHGhpMV1Ei3jTuLCDAhsGYtYdVqgglT9CSgk1IKPzDQZKAEGhn4h3vtVyq6R950rsL6kwm1GqBStjK41kC1yrjwUg3HoRSMcuxxBqQY
foD7nQyUQCMDPw/92g/EnWTAL/1nFwTAr0ZBQKl0/Bw8wNhygcZlVzgIwxF/ApBx+IEOcwGALMC/OLev8WPAc7Yo/Jff9/C1LwXY/5JNRUeZ6xIufaODq98R
DQQyJsXN/wzAz0g8GSiBRgz+xRTVg8C5WxWuvymH3f9u8OJzDBPySHaaxUOBN52rcPqZCsYK/K02owc/kGgyUAIJ/C2Kg4DnARdd2rGhNWJiBEH0n8DfbDOa
8AM9JwMlkMDfUaQAtoBvUtD+RwamA2cMfqDrZKAEEvh7izDao+WyogzCD9QfA2Yafk5asCi1yij8AKAyDb9IlGH4gQFWBkoV/BIdsq2Mww/uNwAI/KK0SOAH
0E8AEPhFWVYK4QeSBgCBX5RCJT4VUgo/kCQApBx+TupPlE2lGH6gVwBIOfyx5DG9qK1SDj/QLQBkAH658os6KgPwA50CgMAvyrIyAj+jXQAQ+EVZVobgBxYG
AIFflGVlDH6gOQAI/KIsK4PwA3EAyDL8jdfbirKl5pM4m/ADgMo0/LFSPcld1F71Y55h+AHAyTT8cuXPsOKWH8WfEpn03jQi8NfP/T7mAiTxJfCLRk+Zg79J
CecCJClQ4BeNnjIJf9P5n2AuQBJfAr9otMQQ+IGecwGS+Bp9+CUmZE+JjnnK4Qe6zgVI4isF8AMol0d+fStRUtVPg8pc/Zh3egCUAfiBjnMBkvhKB/ykgAOv
GAAyKzATqh/kl/eH0dPfdqdGRuAH2s4FSOIrHfAzA65DeOkFg2rVpm4hT1GrlALCkPH8syEcl1pPjwzBD7TMBUjiKx3wx8mOC7yy3+C5Z0JonYbVbkWdxBbQ
mvD8swFefD6E5y0IABmDnxGNBLQLS8kC/LGIgDAAHnmo1qsWohEXIzre/+fBCmpVO38AaAbhZ2arGFSh5jlBSV2lAH4AsAzk84QfPVTDy/sMHKfnvhONoJgB
xwEOHTR48Htl5AsKHLf2Mgg/kQIIFQWgTKQAcBezdMIPALCA1kB5jvHtb5ahVJv7QtHIy1pAKcJ9Xz+CgwfqgR7IJPwAM5ECGGUF5oNECrZzif18TFTEioG/
nmAtUBwjPPyDKh7+QRWeRzCmVwVFoyJjgFyO8H8fruB7355DcUxFfT2ZhD/KFl30cVCBeB9F3d+9d0cK4W8293KEr35xFk/tCZDLSRBIg2L4n33axxc+exCO
W0/IKPzxR4IGMfYphnqWSIOoP9rTBH9chFJRa+BzOw5h96M15HJRL5E8GRg9WRsd01yO8MTuGj7156+gVrPQDh29929ShuAHgZhIgQnPKmJ+klqGwGQL/uai
tAMEAeNzOw7hn/5+DkSA5x0NBPGJJf0EK0fx8YiPDxAdM8chPPAPs9jxZ6+gPGejx35tWnVZgj/u+CAQYOlJx4B2WzZgBkWPRbIJf5zMzNAaYEX45pfn8Ngj
Pq74T0Vsfq3XCAQAYC1LEFghUirq4IsVhoxHH6ni2//7CH68q4ZCQcFxCVbgj83IsgGU3e1oNrvDoGyp3i3YrYS0w9/4Xc87NkZ4dm+Auz91CCef5uC88z2c
cbaL4090UBwjaBk5uCJkDFCeM3h5X4in9vh47EdVPLs3AACMjUcd3Flv9jdvIFIqDKuWmXY71Zzek/N5n+N4G8KwxlRvB2QV/uaNtt4xCAJeei7ET54OoTWQ
LxIKBQXl1M06jSnv4Ji7J3ev59Ffye369EMD2ESOmo94x53SYtP0J6mTeebGMCoVRmXOIgwZrkvI5QlEyHRvf/sNzEp5ZI2/b7wQ7iEAKP33lx/I5SYur9aO
GAK0wN+ajyj6ie81jV1QlwQBoPuB6VXHOlB9kBnd0iSZ5HR0L8Tw98s/1W16fON2bpPkbMp4NGcj1FB0C6DU0WPUyCbwL5Tx3HFdC458Z/s9p1/h1Mt8SCv3
8mgsEPWwT+Y1TfDHrhvuKRpVdvTFkp0cd9mXfRA2/6KfbM5i8oZCjH3908BX/hbv6FrXfq/8jS/U/h1+8fHJ+tj+XgUxMyvlQrF6CAAcANCs/tna8I+x4IgJ
/J2L7XyFSVDPfuGf5zCh3QAHb7jw93bbl7uMv713GPADABGI2YBh/xmozwasWv8HlerhsiJHx54E/gTFCvwCf5sNKxV+gJnI0bXaoTIF4Q8AQJVKrD7yiQ0v
EeMhzy2A0WVUcA+vAn8PO4E/ac4mG4F/gGp0knWdPJjw0Cf+etNLpRIrhcY7AezfKeUmWydE4E/kWOBvzSLw95E0XPjBDFbKASz/XX2TUgCiJ6Qq9/Vq7ZAh
It21KIE/kWOBvzWLwN9H0pDhj5r/0FX/iIF2v17faNX0NNlSidX0X0zsNtb8m+eOEdoOm2jvVeDvYSfwJ83ZZCPwD1CNHmLruGNk2fzb9rtP2V0qsZqeJhu/
CST6S/xFrb32twECfyLHAn9rFoG/j6TFgR+WwVq5ILZfBAA8EDEfjf5nJiLiD//+4RNrTrhHKTVhbYjGqpkCfyLHAn9rFoG/j6RFgp/BrMiBZXuYtHP2J//q
5P0AE0AcRQEi3rmT9fs+ObGfmb+c81YRM0wnrwJ/DzuBP2nOJhuBf4Bq9FDEKgPGc1cRM3/5k3918v6dk6yjsZtNbwWemZkBACgOt/v+nAVICfzJHAv8rVkE
/j6SFhF+ACCQ8oM5y8ZuB4AZzDRyNQWAKVMqsZq+ff0jxtTuy+cmFHPz7GmBv6edwJ80Z5ONwD9ANXroKKsMNjl3lTLGv2/Hl854pFRiNTMz1eB63roAu3ZF
fQKavQ8a4wNNg90F/h52An/SnE02Av8A1eihBaxakDEBQM4HgaOMx5oXAGZmyJRKrEq3r/7XICj/fSG/WjFbI/D3sBP4k+ZsshH4B6hGD82Hn5lN3ptQQTD3
93d8/tR/ja7+NO+1KC1Lg+3aNUMAiLVzYxBUbDRWqIt7gT+xBP7YRuAfoBo9xAvPS1akEJiaVYpuBEB1tuep7XzNnZOsp2bIvO/6l+4aK57wX8uVVwxRm5Xz
BP7EEvhjG4F/gGr0UAv8YLamkD9OV6oH/vKOL2x8z+Qk64VXf6BDACiVoseD/oGXTnSd3C4FtSY0ASjBEkICfxc7gb8fDwL/oPADVpMLBh8kv7pl7bn37AeA
6enplhG+Hd/YsHMn66kpMu/7g33vHSscf2e58mpIRE63ign8XewE/n48CPwDwh+54rCQX+uUaweu2fH5jZ/tdPUHerxeJr4VuOkP9v1jPr/mLdXqQUMdFtEW
+LvYCfz9eBD4jwl+a/Leal0NDv3THZ8/4+e6wQ+06QRs1mNbwACThvueMKgcdpRHzK0ThQT+LnYCfz8eBP5jgh9WK48CUz3sKLwHYNqypXuJPV8w17gVuO7F
3ygWj/9SpXooJIpeJdazvgJ/cqPE2QX+XjY9/aUQ/rrCnDfhVKoH37Hji2fc2+vqD/RoAQDA1BSZUomdD2/fcG+l+sqnxorHO5Zt2LO+An9yo8TZBf5eNj39
pRR+ZhsWcsc51eqBT+344hn3lkrs9IIfSPqKWTDtnIT69tqH1frCxgdyuYnLopYAOe2zC/z9SOBPskng7ww/h/nchFPzZ7/vF/ZdceDA6+zMDGw84aeberYA
IhE/tgV8550XB5XAvN0PKs94btFhbrPYksCf3ChxdoG/l01Pf2mFH2xcp+D4YfUZq8zb77zz4iC67+8NP5C4BRApvqf4k+uevaDgrfkuw06EYc0SRcMFBX6B
vw8PAv+xN/uto3MKRIer4eE3feYL5/5ocnKnbp7s00sJWwCRorkC9zu3bH/Nj2rB7K8o6KrWHrG1VuAX+PvwIPAPAX6tXCKlq74/+yuf+cK5Pypdfr/TD/xA
ny2AWKUSO9PTFN547XNXermJb1gbuKEJrIpbAq31nS+Bv5/SB3HRKELgb92QEvgVkRPU/CNv+/S9Z99Xupyd6e9QmKTkZvXVAog1PU1hqcTOzXeccp9fK79N
kVNznZxq7hMQ+DvYCfwJNgn8HeEHG0fnFJFTC/1yBH9pMPiBAQMA0BwE1t9XCebeSlCHPHdMM3Pnigj8/ZQ+iItGEQJ/64aRh585dJ2iBqlDNTP71jvuPfO+
uDWepOR2GugWoFmN24Hfe+lCL1f4qnZyG8vVg6EiNf8RocDfT+mDuGgUIfC3bhh9+G2Yz61xAlN7OvQrV+/40hmPDNrsb9bALYBYjZbAjvWPHCoffGMQVr43
XjzBYbBhtgn2vsDfriICf+KiuiSOPvzMzACbQv54Jwgr36vNvfDGCP77jxl+YAgtgFjxI8Jrrvmhu76w6TbXK14bBGWY0G//LoG6BP7Wigj8iYvqkpgG+K3R
ytWOW0QYVO6oPvbCH9758MVBkiG+STW0AABE7xGYngYDxDdcv/9drsp9wtG5tdXaoRCAJqI2y48L/P3btBYh8LduGFX4o6s+TM6bcIz1DxhbvX77PRu/ADCV
SqDpaWq/ctcAGmoAiBQNG56aIfOnv/fMWZ63ervnFq+s+bMwJmgMHxb4Wysi8CcuqkviyMMfau06njsOPyjfF5rZ63Z84ewnoqt+suG9/WgRAkCkUul+Z3r6
zSEAvP/6V36HlDPtucV1leoBZoYFYf5tgcDfvwT+thtGEX5mNgSofG4tBWFlH5uw9InPn/ppAIju9998zPf77bRoAQA4+mqx6WmyN/72M6fowuoSmN/jODlV
rR1iEFkCaYG/LxeNIgT+1g2jBz8btqxy3moKTdWC6K6aPTz9mXvOfa6ZnySlDqJFDQCxmlsDN1738qWOdm6Cordp5aFWOwxmhETQjbUIF0jgby1C4G/dMDrw
MzNgCHA8ZwLW+rBsv2E5/ND2e057EAAuv/x+5zuLdNVv1pIEAABgZpqZivoGAOCG3993uVbeH4HtL3ruOPn+ERgbmKifkFRcN4G/tQiBv3XDyoefGYC1bKGU
qz13HEEwxwB9k2146yc/f9p3gOhp2s4ZWBryvX4nLVkAiLWwWfP+6/dfArjvZbZv97zx46wN4AdzYOYQBAKguq1S3E4Cf2sWgb+PpKHBH0HPDCaC4zhjUNqF
78++SkRfAcLPfuJzpz0EtHKxVFryABBrcnKn3rJlkuMvfON/23eStt7VDP51ZntZLjehjQkQhBUwh4YtmAjEDBU9ThT4k7oV+PtIOgb4GcwEsszMRCAiRztO
Hlo5qPpHDEF9H4S/VoH/1Y9/ceMLQAT+rl0z1O8svmFp2QJArFKJ1dZdoKmmgQ03XP/K+ZrUVcx8FbO9xHPHxxU5MNaHMdEPs7VE0e5nBgH19kLH1kJnCfyx
jcDfu8j6U3qAo2Y6gxlEpJRWDrTOQSkXzAZ+MDtLoIcA9S3W+NYn/vKkR+OyJidZb9kCXuor/kItewCIxcy0bdsDetv0Fab5/qd07f6TDTmXWvAbAHsxgHOY
7QbXKSilNJgZzAaGDZht9GOT79M4gvRf38Sl92nTWsSiw899e0gX/AkODAMgUiCo6C/p+meCZYMwrFoCvcjAj4nohwT8S6jDB+/4y9OfbyqFSpc/oLd9Z/45
vpxaMQGgWfX7IQXALoyQf/q7B9cqhU1MwWYAmwnYZJlPJaJ11tq1xCgyuEBEPec5yJU/tpErf68imdmCUAFTmYgOgLEPwE8IeBLQu5WD3bVq8clPfWnNgWbb
bufyStD/BwaH3sOg4P/IAAAAAElFTkSuQmCC
'@

# ========================== РАБОТА С ДАННЫМИ =================================

function Get-DataDirectory {
    # Возвращает папку для хранения конфигурации: рядом со скриптом, если она
    # доступна для записи, иначе %APPDATA%\QuickNotes.
    if ($script:DataDirectory) { return $script:DataDirectory }

    $dir = $PSScriptRoot
    $writable = $false
    if ($dir) {
        try {
            $probe = Join-Path $dir ('.qn-write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
            [System.IO.File]::WriteAllText($probe, 'ok')
            Remove-Item -LiteralPath $probe -Force
            $writable = $true
        } catch {
            $writable = $false
        }
    }
    if (-not $writable) {
        $base = $env:APPDATA
        if (-not $base) { $base = [System.IO.Path]::GetTempPath() }
        $dir = Join-Path $base 'QuickNotes'
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }
    $script:DataDirectory = $dir
    return $dir
}

function Get-ConfigFilePath {
    return (Join-Path (Get-DataDirectory) $script:ConfigName)
}

function Read-ButtonConfig {
    param([string]$Path = (Get-ConfigFilePath))

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    $raw = [System.IO.File]::ReadAllText($Path)
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

    $parsed = $null
    try {
        $parsed = ConvertFrom-Json -InputObject $raw
    } catch {
        # Конфигурация повреждена — убираем её в резервный файл и начинаем с пустого списка.
        try { Move-Item -LiteralPath $Path -Destination ($Path + '.corrupt') -Force } catch { }
        return @()
    }
    if ($null -eq $parsed) { return @() }

    if ($parsed -is [System.Array]) {
        $rawItems = $parsed
    } elseif ($null -ne $parsed.buttons) {
        $rawItems = @($parsed.buttons)
    } else {
        $rawItems = @()
    }

    $result = @()
    foreach ($item in $rawItems) {
        if ($null -eq $item) { continue }
        $name   = [string]$item.name
        $folder = [string]$item.folder
        if ($name) {
            $result += [pscustomobject]@{ Name = $name; Folder = $folder }
        }
    }
    return ,$result
}

function Save-ButtonConfig {
    param(
        [object[]]$Buttons = @(),
        [string]$Path = (Get-ConfigFilePath)
    )

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($b in $Buttons) {
        if ($null -ne $b) {
            $entries.Add([pscustomobject]@{ name = [string]$b.Name; folder = [string]$b.Folder })
        }
    }
    $config = [pscustomobject]@{ buttons = $entries.ToArray() }
    $json   = ConvertTo-Json -InputObject $config -Depth 4

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($true)))
}

# ============================ ВСПОМОГАТЕЛЬНЫЕ ================================

function ConvertTo-SafeFileName {
    # Убирает из имени файла запрещённые символы и крайние точки/пробелы.
    param([string]$Name)

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -notcontains $ch) { [void]$sb.Append($ch) }
    }
    $clean = $sb.ToString().Trim()
    $clean = $clean.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    return $clean
}

function Get-UniqueFilePath {
    # Возвращает несуществующий путь: "имя.txt" -> "имя (2).txt" -> "имя (3).txt" ...
    param(
        [string]$Folder,
        [string]$BaseName,
        [string]$Extension = $script:Extension
    )

    $candidate = Join-Path $Folder ($BaseName + $Extension)
    $i = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $Folder ('{0} ({1}){2}' -f $BaseName, $i, $Extension)
        $i++
    }
    return $candidate
}

function Resolve-FolderPath {
    # Разворачивает переменные окружения и относительные пути.
    param([string]$Path)

    $p = [Environment]::ExpandEnvironmentVariables($Path.Trim())
    if (-not [System.IO.Path]::IsPathRooted($p)) {
        $p = Join-Path (Get-Location).Path $p
    }
    return $p
}

function Open-FolderInExplorer {
    param([string]$Folder)

    if (Test-Path -LiteralPath $Folder -PathType Container) {
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Folder)
    } else {
        Show-MessageBox -Owner $script:Window -Message ("Папка не существует:`n{0}" -f $Folder) -Icon 'Warning'
    }
}

# ================================ РАЗМЕТКА ===================================

$script:CommonStyles = @'
        <Style x:Key="AccentButton" TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#7B6CF6"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#8F83F8"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#6A5BE0"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="#3A3A45"/>
                                <Setter Property="Foreground" Value="#71717C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="GhostButton" TargetType="{x:Type Button}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#C7C7D1"/>
            <Setter Property="BorderBrush" Value="#3A3A45"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#2A2A33"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#4A4A57"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#26262E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DarkTextBox" TargetType="{x:Type TextBox}">
            <Setter Property="Background" Value="#232329"/>
            <Setter Property="Foreground" Value="#ECECF1"/>
            <Setter Property="BorderBrush" Value="#3A3A45"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="CaretBrush" Value="#ECECF1"/>
            <Setter Property="SelectionBrush" Value="#7B6CF6"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#7B6CF6"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="#26262E"/>
                                <Setter Property="Foreground" Value="#71717C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
'@

function Get-MainWindowXaml {
    return (@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Быстрые заметки"
        Width="700" Height="540" MinWidth="540" MinHeight="430"
        WindowStartupLocation="CenterScreen"
        Background="#1B1B1F" Foreground="#ECECF1"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>
[[STYLES]]
        <Style x:Key="NoteButton" TargetType="{x:Type Button}">
            <Setter Property="Background" Value="#2A2A33"/>
            <Setter Property="Foreground" Value="#ECECF1"/>
            <Setter Property="BorderBrush" Value="#3A3A45"/>
            <Setter Property="Width" Value="170"/>
            <Setter Property="Height" Value="92"/>
            <Setter Property="Margin" Value="0,0,12,12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="10">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#33333E"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#7B6CF6"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#3C3C49"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#232329" BorderBrush="#2E2E38" BorderThickness="0,0,0,1" Padding="18,14">
            <DockPanel LastChildFill="True">
                <Button x:Name="BtnAdd" DockPanel.Dock="Right" Style="{StaticResource AccentButton}" Content="+  Добавить кнопку" ToolTip="Создать новую кнопку (Ctrl+N)"/>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="📝" FontSize="22" Margin="0,0,12,0" VerticalAlignment="Center"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Быстрые заметки" FontSize="16" FontWeight="SemiBold"/>
                        <TextBlock Text="Текстовые файлы в один клик" FontSize="11" Foreground="#9A9AA5"/>
                    </StackPanel>
                </StackPanel>
            </DockPanel>
        </Border>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="18,14,6,14">
            <StackPanel>
                <TextBlock x:Name="TxtEmpty" Text="Пока нет ни одной кнопки.&#x0a;Нажмите «+ Добавить кнопку», чтобы создать первую." TextAlignment="Center" Foreground="#9A9AA5" Margin="0,70,0,0" LineHeight="22"/>
                <WrapPanel x:Name="PnlButtons"/>
            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="2" Background="#232329" BorderBrush="#2E2E38" BorderThickness="0,1,0,0" Padding="18,9">
            <DockPanel LastChildFill="True">
                <TextBlock x:Name="TxtCount" DockPanel.Dock="Right" Foreground="#9A9AA5" VerticalAlignment="Center" Margin="18,0,0,0"/>
                <TextBlock x:Name="TxtStatus" Text="Готово к работе" Foreground="#9A9AA5" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}"/>
            </DockPanel>
        </Border>
    </Grid>
</Window>
'@).Replace('[[STYLES]]', $script:CommonStyles)
}

function Get-EditorWindowXaml {
    return (@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Новая заметка"
        Width="720" Height="620" MinWidth="560" MinHeight="480"
        WindowStartupLocation="CenterOwner"
        Background="#1B1B1F" Foreground="#ECECF1"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>
[[STYLES]]
    </Window.Resources>
    <Grid Margin="18,16,18,16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*" MinHeight="140"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <DockPanel Grid.Row="0" LastChildFill="True" VerticalAlignment="Center">
            <TextBlock DockPanel.Dock="Left" Text="Имя:" Foreground="#C7C7D1" Margin="0,0,10,0" VerticalAlignment="Center"/>
            <TextBlock DockPanel.Dock="Right" Text=".txt" Foreground="#71717C" Margin="10,0,0,0" FontFamily="Consolas" VerticalAlignment="Center"/>
            <TextBox x:Name="TbFileName" Style="{StaticResource DarkTextBox}" MaxLength="180"/>
        </DockPanel>

        <TextBlock x:Name="TxtFolder" Grid.Row="1" Margin="2,9,2,0" Foreground="#9A9AA5" FontSize="11" TextTrimming="CharacterEllipsis"/>

        <TextBox x:Name="TbText" Grid.Row="2" Style="{StaticResource DarkTextBox}" Margin="0,10,0,0"
                 AcceptsReturn="True" AcceptsTab="True" TextWrapping="Wrap" VerticalContentAlignment="Stretch"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                 FontFamily="Consolas" FontSize="14" Padding="12,10"/>

        <DockPanel Grid.Row="3" LastChildFill="False" Margin="2,14,2,0">
            <TextBlock DockPanel.Dock="Left" Text="Ctrl+S — сохранить   •   Esc — закрыть" Foreground="#71717C" FontSize="11" VerticalAlignment="Center"/>
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="BtnCancel" Style="{StaticResource GhostButton}" Content="Отмена" MinWidth="98" IsCancel="True"/>
                <Button x:Name="BtnSave" Style="{StaticResource AccentButton}" Content="Сохранить" MinWidth="118" Margin="10,0,0,0" ToolTip="Ctrl+S"/>
            </StackPanel>
        </DockPanel>
    </Grid>
</Window>
'@).Replace('[[STYLES]]', $script:CommonStyles)
}

function Get-ButtonDialogXaml {
    return (@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Новая кнопка"
        Width="520" Height="300" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner"
        Background="#1B1B1F" Foreground="#ECECF1"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>
[[STYLES]]
    </Window.Resources>
    <Grid Margin="20,18,20,18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Название кнопки:" Foreground="#C7C7D1"/>
        <TextBox Grid.Row="1" x:Name="TbName" Style="{StaticResource DarkTextBox}" Margin="0,8,0,0" MaxLength="60"/>

        <TextBlock Grid.Row="2" Text="Папка, в которую будут сохраняться .txt:" Foreground="#C7C7D1" Margin="0,16,0,0"/>
        <DockPanel Grid.Row="3" LastChildFill="True" Margin="0,8,0,0">
            <Button DockPanel.Dock="Right" x:Name="BtnBrowse" Style="{StaticResource GhostButton}" Content="Обзор…" MinWidth="96" Margin="10,0,0,0"/>
            <TextBox x:Name="TbFolder" Style="{StaticResource DarkTextBox}"/>
        </DockPanel>

        <DockPanel Grid.Row="4" LastChildFill="False" Margin="0,18,0,0">
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="BtnCancel" Style="{StaticResource GhostButton}" Content="Отмена" MinWidth="98" IsCancel="True"/>
                <Button x:Name="BtnSave" Style="{StaticResource AccentButton}" Content="Сохранить" MinWidth="118" Margin="10,0,0,0"/>
            </StackPanel>
        </DockPanel>

        <TextBlock Grid.Row="5" x:Name="TxtError" Foreground="#F2777A" TextWrapping="Wrap" Margin="0,12,0,0" Visibility="Collapsed"/>
    </Grid>
</Window>
'@).Replace('[[STYLES]]', $script:CommonStyles)
}

# ============================== ИНТЕРФЕЙС ====================================

function ConvertFrom-XamlMarkup {
    param([string]$Markup)
    return [System.Windows.Markup.XamlReader]::Parse($Markup)
}

function Get-AppIconImageSource {
    # Декодирует встроенную иконку; при неудаче пробует файл AppIcon.ico рядом
    # со скриптом. Возвращает $null, если иконку получить не удалось.
    $candidates = @()
    try {
        $b64 = $script:Base64Icon -replace '\s', ''
        if ($b64) { $candidates += ,([System.Convert]::FromBase64String($b64)) }
    } catch { }

    $iconFile = $null
    if ($PSScriptRoot) {
        $p = Join-Path $PSScriptRoot 'AppIcon.ico'
        if (Test-Path -LiteralPath $p -PathType Leaf) { $iconFile = $p }
    }

    $loaded = $null
    try {
        Add-Type -AssemblyName PresentationCore -ErrorAction Stop
        foreach ($bytes in $candidates) {
            $ms = New-Object System.IO.MemoryStream
            try {
                $ms.Write($bytes, 0, $bytes.Length)
                $ms.Position = 0
                $loaded = [System.Windows.Media.Imaging.BitmapFrame]::Create(
                    $ms,
                    [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
                    [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
                break
            } catch {
                $loaded = $null
            } finally {
                $ms.Dispose()
            }
        }
        if (-not $loaded -and $iconFile) {
            $fs = [System.IO.File]::OpenRead($iconFile)
            try {
                $loaded = [System.Windows.Media.Imaging.BitmapFrame]::Create(
                    $fs,
                    [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
                    [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
            } finally {
                $fs.Dispose()
            }
        }
    } catch {
        return $null
    }
    return $loaded
}

function Show-MessageBox {
    param(
        [System.Windows.Window]$Owner,
        [string]$Message,
        [string]$Title = $script:AppTitle,
        [string]$ButtonsText = 'OK',
        [string]$IconText = 'Information'
    )

    $button = [System.Windows.MessageBoxButton]::($ButtonsText)
    $icon   = [System.Windows.MessageBoxImage]::($IconText)
    if ($Owner) {
        return [System.Windows.MessageBox]::Show($Owner, $Message, $Title, $button, $icon)
    }
    return [System.Windows.MessageBox]::Show($Message, $Title, $button, $icon)
}

function Update-Status {
    param([string]$Text, [switch]$Good)

    if ($script:TxtStatus) {
        $script:TxtStatus.Text = $Text
        if ($Good) { $script:TxtStatus.Foreground = '#7DD487' }
        else       { $script:TxtStatus.Foreground = '#9A9AA5' }
    }
}

function Refresh-Buttons {
    $script:PnlButtons.Children.Clear()

    $buttons = @($script:Buttons)
    $script:TxtCount.Text = ('Кнопок: {0}' -f $buttons.Count)

    if ($buttons.Count -eq 0) {
        $script:TxtEmpty.Visibility   = 'Visible'
        $script:PnlButtons.Visibility = 'Collapsed'
    } else {
        $script:TxtEmpty.Visibility   = 'Collapsed'
        $script:PnlButtons.Visibility = 'Visible'
    }

    foreach ($b in $buttons) {
        if ($null -eq $b) { continue }
        [void]$script:PnlButtons.Children.Add((New-ButtonCard -Button $b))
    }
}

function New-ButtonCard {
    param([object]$Button)

    $btn = New-Object System.Windows.Controls.Button
    $btn.Style = $script:Window.TryFindResource('NoteButton')
    $btn.Tag   = $Button
    $btn.ToolTip = ("{0}`n`nЛевый клик — создать заметку.`nПравый клик — меню." -f $Button.Folder)

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.Text = '📝'
    $icon.FontSize = 24
    $icon.HorizontalAlignment = 'Center'

    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = $Button.Name
    $label.TextWrapping = 'Wrap'
    $label.TextAlignment = 'Center'
    $label.TextTrimming = 'CharacterEllipsis'
    $label.FontWeight = 'SemiBold'
    $label.Margin = '10,7,10,0'
    $label.MaxWidth = 140
    $label.MaxHeight = 42

    $stack = New-Object System.Windows.Controls.StackPanel
    $stack.VerticalAlignment = 'Center'
    [void]$stack.Children.Add($icon)
    [void]$stack.Children.Add($label)
    $btn.Content = $stack

    # Левый клик — открыть редактор заметки.
    $btn.Add_Click({
        param($sender, $e)
        Open-Editor -ButtonObj $sender.Tag
    })

    # Правый клик — контекстное меню.
    $menu = New-Object System.Windows.Controls.ContextMenu

    $miEdit = New-Object System.Windows.Controls.MenuItem
    $miEdit.Header = '✏  Изменить'
    $miEdit.Tag = $Button
    $miEdit.Add_Click({
        param($sender, $e)
        Show-ButtonDialog -Existing $sender.Tag
    })
    [void]$menu.Items.Add($miEdit)

    $miDelete = New-Object System.Windows.Controls.MenuItem
    $miDelete.Header = '🗑  Удалить'
    $miDelete.Tag = $Button
    $miDelete.Add_Click({
        param($sender, $e)
        $target = $sender.Tag
        $answer = Show-MessageBox -Owner $script:Window -Message ("Удалить кнопку «{0}»?`n(Уже созданные файлы останутся в папке.)" -f $target.Name) -Title 'Удаление кнопки' -ButtonsText 'YesNo' -IconText 'Question'
        if ($answer -ne 'Yes') { return }
        $script:Buttons = @($script:Buttons | Where-Object { $_ -ne $target })
        try {
            Save-ButtonConfig -Buttons $script:Buttons
        } catch {
            Show-MessageBox -Owner $script:Window -Message ("Не удалось сохранить настройки:`n{0}" -f $_.Exception.Message) -Icon 'Error'
            return
        }
        Refresh-Buttons
        Update-Status ('Кнопка «{0}» удалена' -f $target.Name)
    })
    [void]$menu.Items.Add($miDelete)

    [void]$menu.Items.Add((New-Object System.Windows.Controls.Separator))

    $miOpen = New-Object System.Windows.Controls.MenuItem
    $miOpen.Header = '📂  Открыть папку'
    $miOpen.Tag = $Button
    $miOpen.Add_Click({
        param($sender, $e)
        Open-FolderInExplorer -Folder (Resolve-FolderPath $sender.Tag.Folder)
    })
    [void]$menu.Items.Add($miOpen)

    $btn.ContextMenu = $menu
    return $btn
}

function Show-ButtonDialog {
    # Диалог создания/редактирования пользовательской кнопки.
    param([object]$Existing = $null)

    $dlg = ConvertFrom-XamlMarkup (Get-ButtonDialogXaml)
    $dlg.Owner = $script:Window
    if ($script:AppIconImageSource) { $dlg.Icon = $script:AppIconImageSource }
    $tbName    = $dlg.FindName('TbName')
    $tbFolder  = $dlg.FindName('TbFolder')
    $btnBrowse = $dlg.FindName('BtnBrowse')
    $btnSave   = $dlg.FindName('BtnSave')
    $btnCancel = $dlg.FindName('BtnCancel')
    $txtError  = $dlg.FindName('TxtError')

    if ($null -ne $Existing) {
        $dlg.Title     = 'Изменить кнопку'
        $tbName.Text   = $Existing.Name
        $tbFolder.Text = $Existing.Folder
    } else {
        $dlg.Title = 'Новая кнопка'
    }

    $btnBrowse.Add_Click({
        param($sender, $e)
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'Выберите папку, в которую будут сохраняться текстовые файлы'
        $fbd.ShowNewFolderButton = $true
        $current = [Environment]::ExpandEnvironmentVariables($tbFolder.Text.Trim())
        if ($current -and (Test-Path -LiteralPath $current -PathType Container)) {
            try { $fbd.SelectedPath = (Resolve-Path -LiteralPath $current).Path } catch { }
        }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tbFolder.Text = $fbd.SelectedPath
        }
    })

    $btnCancel.Add_Click({
        param($sender, $e)
        $dlg.Close()
    })

    $saveAction = {
        $name   = $tbName.Text.Trim()
        $folder = [Environment]::ExpandEnvironmentVariables($tbFolder.Text.Trim())

        if (-not $name) {
            $txtError.Text = 'Введите название кнопки.'
            $txtError.Visibility = 'Visible'
            return
        }
        if (-not $folder) {
            $txtError.Text = 'Укажите папку для файлов.'
            $txtError.Visibility = 'Visible'
            return
        }
        if (-not [System.IO.Path]::IsPathRooted($folder)) {
            $folder = Join-Path (Get-Location).Path $folder
        }

        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            $answer = Show-MessageBox -Owner $dlg -Message ("Папка не существует:`n{0}`n`nСоздать её?" -f $folder) -ButtonsText 'YesNo' -IconText 'Question'
            if ($answer -ne 'Yes') { return }
            try {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
            } catch {
                Show-MessageBox -Owner $dlg -Message ("Не удалось создать папку:`n{0}" -f $_.Exception.Message) -Icon 'Error'
                return
            }
        }

        if ($null -ne $Existing) {
            $Existing.Name   = $name
            $Existing.Folder = $folder
        } else {
            $script:Buttons = @($script:Buttons) + [pscustomobject]@{ Name = $name; Folder = $folder }
        }

        try {
            Save-ButtonConfig -Buttons $script:Buttons
        } catch {
            Show-MessageBox -Owner $dlg -Message ("Не удалось сохранить настройки:`n{0}" -f $_.Exception.Message) -Icon 'Error'
            return
        }

        $dlg.DialogResult = $true
        $dlg.Close()
    }

    $btnSave.Add_Click($saveAction)
    $tbName.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $e.Handled = $true
            & $saveAction
        }
    })
    $tbFolder.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $e.Handled = $true
            & $saveAction
        }
    })

    [void]$dlg.ShowDialog()
    Refresh-Buttons
}

function Open-Editor {
    # Окно заметки: имя по умолчанию — текущая дата, текст редактируется, файл
    # создаётся по кнопке «Сохранить».
    param([object]$ButtonObj)

    $folder = Resolve-FolderPath $ButtonObj.Folder
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        $answer = Show-MessageBox -Owner $script:Window -Message ("Папка не существует:`n{0}`n`nСоздать её?" -f $folder) -ButtonsText 'YesNo' -IconText 'Question'
        if ($answer -ne 'Yes') { return }
        try {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        } catch {
            Show-MessageBox -Owner $script:Window -Message ("Не удалось создать папку:`n{0}" -f $_.Exception.Message) -Icon 'Error'
            return
        }
    }

    # Имя по умолчанию — текущая дата (с номером, если файл уже есть).
    $today = Get-Date -Format $script:DateFormat
    $suggested = [System.IO.Path]::GetFileNameWithoutExtension((Get-UniqueFilePath -Folder $folder -BaseName $today))

    $dlg = ConvertFrom-XamlMarkup (Get-EditorWindowXaml)
    $dlg.Title = ('Новая заметка — {0}' -f $ButtonObj.Name)
    $dlg.Owner = $script:Window
    if ($script:AppIconImageSource) { $dlg.Icon = $script:AppIconImageSource }

    $tbFileName = $dlg.FindName('TbFileName')
    $tbText     = $dlg.FindName('TbText')
    $txtFolder  = $dlg.FindName('TxtFolder')
    $btnSave    = $dlg.FindName('BtnSave')
    $btnCancel  = $dlg.FindName('BtnCancel')

    $tbFileName.Text = $suggested
    $txtFolder.Text  = ('Папка: {0}' -f $folder)
    $txtFolder.ToolTip = $folder

    $saveAction = {
        Save-FromEditor -Editor $dlg -Folder $folder -TbName $tbFileName -TbText $tbText
    }

    $btnSave.Add_Click($saveAction)
    $btnCancel.Add_Click({
        param($sender, $e)
        $dlg.Close()
    })

    $dlg.Add_PreviewKeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::S -and
            (($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0)) {
            $e.Handled = $true
            & $saveAction
        }
    })

    # Enter в поле имени — перейти к тексту.
    $tbFileName.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
            $e.Handled = $true
            $tbText.Focus()
        }
    })

    $dlg.Add_ContentRendered({
        param($sender, $e)
        $tbFileName.Focus()
        $tbFileName.SelectAll()
    })

    $result = $dlg.ShowDialog()
    if ($result -and $script:LastSavedPath) {
        Update-Status ('✔ Сохранено: {0}' -f $script:LastSavedPath) -Good
    }
}

function Save-FromEditor {
    param(
        [System.Windows.Window]$Editor,
        [string]$Folder,
        [object]$TbName,
        [object]$TbText
    )

    $name = ConvertTo-SafeFileName $TbName.Text
    if (-not $name) {
        Show-MessageBox -Owner $Editor -Message 'Введите имя файла.' -Icon 'Warning'
        return
    }

    $path = Join-Path $Folder ($name + $script:Extension)

    if (Test-Path -LiteralPath $path) {
        $message = ("В папке уже есть файл:`n{0}{1}`n`nДа — перезаписать его.`nНет — сохранить с номером в имени.`nОтмена — вернуться к редактированию." -f $name, $script:Extension)
        $answer = Show-MessageBox -Owner $Editor -Message $message -Title 'Файл уже существует' -ButtonsText 'YesNoCancel' -IconText 'Warning'
        if ($answer -eq 'Cancel') { return }
        if ($answer -eq 'No') {
            $path = Get-UniqueFilePath -Folder $Folder -BaseName $name
        }
    }

    try {
        $encoding = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($path, $TbText.Text, $encoding)
    } catch {
        Show-MessageBox -Owner $Editor -Message ("Не удалось сохранить файл:`n{0}" -f $_.Exception.Message) -Icon 'Error'
        return
    }

    $script:LastSavedPath = $path
    $Editor.DialogResult = $true
    $Editor.Close()
}

# ============================== САМОТЕСТЫ ====================================

function Invoke-SelfTest {
    function Assert-True {
        param([bool]$Condition, [string]$Message)
        if (-not $Condition) { throw $Message }
    }

    $passed = 0
    $failed = 0
    $failures = @()

    $checks = @(
        @{ Name = 'safe-name: removes forbidden slash'; Test = { (ConvertTo-SafeFileName 'note/test') -eq 'notetest' } },
        @{ Name = 'safe-name: trims trailing dots/spaces'; Test = { (ConvertTo-SafeFileName '  report. ') -eq 'report' } },
        @{ Name = 'safe-name: null for empty result'; Test = { $null -eq (ConvertTo-SafeFileName '///') } },
        @{ Name = 'safe-name: keeps normal text'; Test = { (ConvertTo-SafeFileName 'Встреча 2026') -eq 'Встреча 2026' } },
        @{ Name = 'date format is yyyy-MM-dd'; Test = { (Get-Date -Format $script:DateFormat) -match '^\d{4}-\d{2}-\d{2}$' } },
        @{ Name = 'unique-file: appends (2)'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $tmp | Out-Null
            try {
                Set-Content -LiteralPath (Join-Path $tmp '2026-01-01.txt') -Value 'x'
                $p = Get-UniqueFilePath -Folder $tmp -BaseName '2026-01-01'
                ([System.IO.Path]::GetFileName($p) -eq '2026-01-01 (2).txt')
            } finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
        } },
        @{ Name = 'unique-file: free name returned as-is'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $tmp | Out-Null
            try {
                $p = Get-UniqueFilePath -Folder $tmp -BaseName 'free'
                ([System.IO.Path]::GetFileName($p) -eq 'free.txt')
            } finally { Remove-Item -LiteralPath $tmp -Recurse -Force }
        } },
        @{ Name = 'config: json round-trip'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N') + '.json')
            try {
                $src = @(
                    [pscustomobject]@{ Name = 'Работа'; Folder = 'C:\Notes' },
                    [pscustomobject]@{ Name = 'Дом';   Folder = 'D:\Личное' }
                )
                Save-ButtonConfig -Buttons $src -Path $tmp
                $restored = @(Read-ButtonConfig -Path $tmp)
                Assert-True ($restored.Count -eq 2) 'expected 2 buttons'
                Assert-True ($restored[0].Name -eq 'Работа' -and $restored[0].Folder -eq 'C:\Notes') 'first button mismatch'
                Assert-True ($restored[1].Name -eq 'Дом' -and $restored[1].Folder -eq 'D:\Личное') 'second button mismatch'
                $true
            } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } }
        } },
        @{ Name = 'config: empty list round-trip'; Test = {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('qn-test-' + [guid]::NewGuid().ToString('N') + '.json')
            try {
                Save-ButtonConfig -Buttons @() -Path $tmp
                $restored = @(Read-ButtonConfig -Path $tmp)
                Assert-True ($restored.Count -eq 0) 'expected empty list'
                $true
            } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } }
        } },
        @{ Name = 'config: missing file returns empty'; Test = {
            $restored = @(Read-ButtonConfig -Path (Join-Path ([System.IO.Path]::GetTempPath()) ('qn-missing-' + [guid]::NewGuid().ToString('N') + '.json')))
            ($restored.Count -eq 0)
        } },
        @{ Name = 'embedded icon is valid base64 ICO'; Test = {
            $b64 = $script:Base64Icon -replace '\s', ''
            Assert-True ($b64.Length -gt 1000) 'icon base64 suspiciously short'
            $bytes = [System.Convert]::FromBase64String($b64)
            Assert-True ($bytes.Length -gt 1000) 'decoded icon is too small'
            Assert-True (($bytes[0] -eq 0) -and ($bytes[1] -eq 0) -and ($bytes[2] -eq 1) -and ($bytes[3] -eq 0)) 'decoded data is not an ICO file'
            $true
        } },
        @{ Name = 'main XAML is well-formed'; Test = {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-MainWindowXaml))
            $true
        } },
        @{ Name = 'editor XAML is well-formed'; Test = {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-EditorWindowXaml))
            $true
        } },
        @{ Name = 'button dialog XAML is well-formed'; Test = {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-ButtonDialogXaml))
            $true
        } }
    )

    Write-Host 'QuickNotes self-test:'
    foreach ($check in $checks) {
        $ok = $false
        $err = $null
        try { $ok = [bool](& $check.Test) } catch { $err = $_.Exception.Message }
        if ($ok) {
            $passed++
            Write-Host ('  [OK]   ' + $check.Name)
        } else {
            $failed++
            if (-not $err) { $err = 'condition is false' }
            $failures += ('{0}: {1}' -f $check.Name, $err)
            Write-Host ('  [FAIL] ' + $check.Name + ' :: ' + $err)
        }
    }

    Write-Host ''
    Write-Host ('Passed: {0}, failed: {1}' -f $passed, $failed)
    if ($failed -gt 0) {
        throw ('Self-test failed: ' + ($failures -join '; '))
    }
}

# ================================ ЗАПУСК =====================================

if ($SelfTest) {
    try {
        Invoke-SelfTest
        exit 0
    } catch {
        Write-Error $_
        exit 1
    }
}

try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms

    # Своя идентичность в панели задач (AppUserModelID). Без этого окно остаётся
    # «привязано» к powershell.exe: кнопка на панели задач группируется с ним и
    # может показывать значок PowerShell вместо иконки приложения.
    try {
        Add-Type -Namespace Win32 -Name ShellApi -MemberDefinition @'
[DllImport("shell32.dll", SetLastError = true)]
public static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);
'@ -ErrorAction SilentlyContinue
        [void][Win32.ShellApi]::SetCurrentProcessExplicitAppUserModelID('SweetBaget.QuickNotes')
    } catch {
        # Не критично: останется системная группировка по процессу.
    }

    $script:DataDirectory  = $null
    $script:LastSavedPath  = $null
    $script:Buttons        = @(Read-ButtonConfig)
    $script:FirstRun       = -not (Test-Path -LiteralPath (Get-ConfigFilePath))

    $window = ConvertFrom-XamlMarkup (Get-MainWindowXaml)
    $script:AppIconImageSource = Get-AppIconImageSource
    if ($script:AppIconImageSource) { $window.Icon = $script:AppIconImageSource }
    $script:Window     = $window
    $script:BtnAdd     = $window.FindName('BtnAdd')
    $script:TxtEmpty   = $window.FindName('TxtEmpty')
    $script:PnlButtons = $window.FindName('PnlButtons')
    $script:TxtStatus  = $window.FindName('TxtStatus')
    $script:TxtCount   = $window.FindName('TxtCount')

    Update-Status ('Настройки: {0}' -f (Get-ConfigFilePath))

    $script:BtnAdd.Add_Click({
        param($sender, $e)
        Show-ButtonDialog
    })

    $window.Add_PreviewKeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::N -and
            (($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0)) {
            $e.Handled = $true
            Show-ButtonDialog
        }
    })

    # При первом запуске сразу предлагаем создать кнопку.
    $window.Add_ContentRendered({
        param($sender, $e)
        if ($script:FirstRun) {
            $script:FirstRun = $false
            Show-ButtonDialog
        }
    })

    Refresh-Buttons
    [void]$window.ShowDialog()
} catch {
    $message = ('Ошибка: {0}' -f $_.Exception.Message)
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [void][System.Windows.MessageBox]::Show($message, $script:AppTitle, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } catch {
        Write-Error $message
    }
    exit 1
}
