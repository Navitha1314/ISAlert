#####################################################################################################
### ENV    : PRODUCTION
### AUTHOR : NAVITHA GOVINDRAJ
### DATE   : 04 JULY 2024
### REASON : SCRIPT TO RECEIVE INPUT FROM SM FOR IS ALERT (LOCATION AND ALERT MESSAGE)
#####################################################################################################
$logfile = "$env:SystemDrive\OHMSILOGS\ReceiveInputforISAlert.log"
Start-Transcript -Path $logfile -Append -Force
Add-Type -AssemblyName system.windows.Forms
$Form = New-Object system.windows.forms.form
$Form.Text = "OHIOHEALTH - IS ALERT MANAGEMENT SYSTEM" 
$Font1 = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Regular)
$Font2 = New-Object System.Drawing.Font("Cambria",14,[System.Drawing.FontStyle]::Regular)
$form.Height = '600'
$form.Width = '650'
$form.AutoSize = $True
$Form.AutoScale = $True
$Form.BackColor = "White"
$Form.Font = $Font1
$iconBase64 = '/9j/4AAQSkZJRgABAQEAkACQAAD/4QAiRXhpZgAATU0AKgAAAAgAAQESAAMAAAABAAEAAAAAAAD/2wBDAAIBAQIBAQICAgICAgICAwUDAwMDAwYEBAMFBwYHBwcGBwcICQsJCAgKCAcHCg0KCgsMDAwMBwkODw0MDgsMDAz/2wBDAQICAgMDAwYDAwYMCAcIDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAz/wAARCAAdAB0DASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDJ/wCEqb+8fz/+vWtp9hrWq+H11a30+4k0yS9j05LosI4nuZCqpECzDLEsoOOF3DJAIry3wF4m0ufx1pK619ufR1n8y8Wyh8+4kjQF9iJhslioU5UgBiSCARX0x8Tf2mvCXwK1zw9/wi91b6t5GvQ6tdaVZ2cMGlRIGe2lwlnbWrRzFYI5VRXAcOjsG8xs/wBH5hUqUKkaVGHM3rs7el+5/J2V0qWIpyrV6vLGOnS+vWz6em55j4il1LwjqTWeqWd1p90oU+XMhXIZVcEHOGBV1bKkjDA9xVH/AISpv7x/P/69ehfH+8s/G3wd/wCEo1TTfHVhDHCr+HJJrIRWxYWdvbzedGsLOoebTLhdzyqFEathQw3/ADYvizj7361tlr+s0ueSs1o+1/6+56HPm0Xg63JGV4vVPrbz/TutTkvBnxF1n4b+MdN1/RrtrPVNJmE9tMBnawyCCOhUglSDwQxB4Jr6q139qnwL8Uf2cNQtfEGuTah4i1XUNLnuILqEQTW2oGzt45rpFSSGP7PFNHOFc+cFTyzJE5PHx6RmmeQvvXr5hk9LFTjUd4yjs1vve3oeNlPEFfAU50YpShPdO9r2tf1PpH9qH9sGPWLDxB4R0G8bWrfUGgDa7EtjDFJC0ETXFuIorGNyfOEi+as4Up8u1h8x+d01Gbb96q6RBKdXRl+W0cJS9nSXr5vv/wANocmaZtXx9b21Z+SXZXvb/gvU/9k='
$iconBytes = [Convert]::FromBase64String($iconBase64)
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
$Form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))
$banner_base64 = 'iVBORw0KGgoAAAANSUhEUgAAAl4AAACfCAIAAACnaEDpAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFiUAABYlAUlSJPAAACBrSURBVHhe7Z3NixzHGcb1l+iYa0DHxH+ByMXXnIMPS2BhhmG8+LAQORc50SH2xQgLQ2SwEIhhZtZkJSxFVmKtsQyWiBHeEEvIsQVGxsjIxMExzvPWW9vbXV1VXd3z1dPz/HgQq+mq6urqqnqqqr9O/EQIIYSQHLRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArRGQgghpACtkRBCCClAaySEEEIK0BoJIYSQArTGIJ98+XR874sLNw6/+PZ7+1MVH3z+zSvvHl69+2/8YX8ihBCybtAaXWCEp87dPDGYnOiP9V94pN1WxaU7j05sjyQi1Bv//LW/2Q2EEELWB1rjT0//+z/7l+Hh189O7OydeOkdEf7Y2Uu3RswyxRSzuLv7doMBO0qfgBJCCFkVG22NH3z+zQtv3cE877sffrQ/qTWqsc1ojdCZa3aDAekgteffOLj66Vf2J0IIIe1jQ63x0p1HdtUUWq41nhhOsdOf/f7axYMHdgMhhJA2saHWeGLriliUGtiSrTFLeXffWcslhBDSBjbVGrdHxwa2Emt86R1MHGmNhBDSQmiNbbHGwyfPeA2SEELaAK2xLdZ48OibE1tXti5/zKkkIYSsFlpjW6zxg8+/OdGTJynx+1//9cT+SgghZOnQGltmjbrTweSVdw/tBkIIIcuF1tg+a1Tt7HFllRBCVgKtsZXWiDB8tIMQQlYErZHWSAghpACtcd7WOJgcp0xrJISQNYTWWM8ar97996lzN5977ZZXJ89ez5xPRGskhJA1hNZYzxovvf+ZGNhw6teR7VnRGgkhZA2hNda3xvzVxLhojYQQsobQGmmNhBBCCtAa61sj4jrrqJmObM+K1kgIIWvIGlvjmwcPxY3US6L6xR9v5M0PNLZG/P2H6T9Ckg8jHzmfiNZICCFryLpao7yMWx+TULOJaDidozXG4XONhBDSAdbSGuEZ8o3+tCkjbIbWSAghJJ21tMatyx8XnqyPi9ZICCGkDutnjVc//erE1hXxNhhJirZHMCHXGn9z+TjAby671mhi2X+3R8uxRlkixu5g+Sq+XpwQQlbE+lnjF99+D/eqKxv5iFpb88Z56/7jX/3pvV+//nevMD3NnK+uNR4+eYbZ8Iuje5kcOyeEELIc1vU2nFWxuLfhEEIIaQm0xnos7pF/QgghLYHWWA9aIyGEdJ51ssbvfvhxJbK7N4g1bl05vlnG0bD4PAmtkRBC1pC1scZX3j08efb6qXM3l6/DJ8f36cDeXr3xT6/ePHgoT5Xk3bHmbTiDtz96eXS3mZABmxAhhJDZWA9rhIWc2N23vrJMmZtrYGk2H1XM8+GNutoeIQWbVufAuOGv/3py9dOvMP6AMA5wXtEHXbhxaEMTsni+++HHrFrqyNhbLS+9/5mNsDDQjaBvQTYu3XmkDcTJA4TR88PSrfgkwnpY48WDBzWu8M1RxsmWY42FR/4baDh94a07Nq0OgXYuK9goGQhlq3KGBVB/jPK0cZYCBiLaFaI/+uLb7+2vZDMYvP1RYrX89et/t3EWACoeOhkZUms2zB4LGci0PYJ32mgkgfWwRvfr+UvTGlmjSbl7fbRYY75IQ9rZ++Ur122cBYNCfu61W8dTfGSvN37lXc5ZNwixRufGAq+G00VbI7oRd6de9caY4NpoJIE1sEZZadQBGvqgJQs7XebbcHCYZgm3obZHq73iiEPAFOrS+5+9OLq3dflj6MKNQ/wXxZK/XluLtlnjdz/8KGv7KG0nA9uj3/3lvg1Eus4SrBFj5cqeh9a4ONbAGnH60eeuUM5NqhFmsUaYxwtv3VFHaazlWyPODtzr1Lmb4us49mwBR91aJ1WQWfM5ff42vDO9PEHbrBH+F8xPb5w+iloQmLyKc5+5Vk+7+8+/cWCTKPHw62dSaZ0oKXrpna52xwuyRtjhxYMHaCZSx7auVC5F0BoXx3osqK4Ls1jj2oHDQX8qE2v439FBVQjBUD6DCVwcrdomFAVmb5cu0RPlSs/VUqxRpozo8UN5GExWvqwac+6Ioj04Bm1yCpwoKeqNu3p967d//tDWfK2WzoFnSrBGdALWDnUcmbWmwaRypItGJCMhxMpaRygztMaa0BrnyYZYIxqkfLRZG2R2sOlCCaD9pxkJCk3vdsEen3vtVvCq89KsET2Rs+tMg8mLo3s26IoQa2xwUnb2IrPGR0//Q2t0gM2gWr48uqvVUqfIHiVY48Ejc5NBrmewSrBGVMhLdx6hdWAWizMoizd6J38+HRWtsSa0xnmyCdYox4jmB2/LDrOZUA79MbqVuoXgb/zLnDU6u87UklkjrXEV+CtGojV6T1mCNZZBgRe6oEy0xpq03RrRb65cNisJdN4a0VbFFMvO1FiDCcohcXFVkVgrskYg0+XQsKC/+muNYo06BXHyFtfcF1S1nm/SAwO0xo7RamvU9St0natV+t2Vs1jjB59/g1/Qv88oJLKgDtpex6rb7VZqOEWe091RopTzsCxrFJ/wFkJ/DNe0gVYHekb0xbLsjN4WfhY/XzBRCGF29yNLwTg1dqUOCeoFNiedTNgpAhgffe61W7/984fpbWfdoTV2jFZbozzq5z3NSxO6lSW/DUcDzKL+Qp6xQ5oV5wK7RgtHv4lgWZ+bOMUcTiMLeg4SvpzmsqwRSO9jylmOV9UbI//tWQNATuBnD79+hqyiWPynoDdGlUOdRMjKnGOQqmkiitwz7E3QnEQE0AQRxUbeDGiNHaPV1ijv8o4MUZcgdAHr8sh/puh1o2bIccVNDltNLwAHRePEkaIdIhb+ix9lwuFtrnn1x4kdgQQu52SJ1gjQ+2MOjUPD3AiTRRyp3dA+fvWn9/wnbntkQ9Qk2JUPp1uXP7aBNg9aY8dotTW+OLrnrzRLk3GytbNGmMQcx+ywAf+dL5nMWmJk6QwpoJFLIlVnM+U1jxKynJnlWmNGy+dGyN4v/njDf+5ojXOF1tgxWm2NcoUDpxkzkhQ5VaHW1pCw9zrPcV+680gWRVHRIVNB7QaDWKMejm7d3c9bo7SQLfO5K43eWGa/c3yVMOagkiunAFXY3e5+4pwJWZLb3ENJQWkPP0jI1lhjy6E1Lg1aY8dotTWiz5UXjyVI7hDJ97mDyZsHD7OtMvss9sj5rZXKG1gcmB9yksm55qercPmt+TkHnAP5zAeYRdiXTXc20Jz8LQ1Ch7u7nz5uADhecUdvR6Da3a/MuQSjNaZBa3RAW84r3wBnhNbYMVptjelI9cqvRm6P8pW+UF3QTezsoVXYbSSKOJm3YzXF2KCx2eVZJ7VM/TFGLTZoAAmWZo04yxhdYaDw2z9/iB4Kwh8Yf9Qa7iioTtqZqjAgQJVz1LifRcTDJ89w4Mjb4O2PTp+/rbnF38g8cothU7PEEWstrBH5RJHCCXDIL7x1J3+yUCyND1/PFFLQOvD8Gwcnz17P69S5m1rOCINT0GwvSro15isSJL1TwBqRbScwhBYUyWe6NXrLHH9jyI7fZymKbtARa5TJTdEaUYfsNp81zmtS1W3QQvzNDJrh8Xa5Ihu6qjqc/upP79lwASRYua8fTuEoNsTRC2mlx8GO8G9eg4n8uLOHbreyGqCD+MP0H0hZ1va179seiZCCo60rlaZeBhlAMUriSAFFjbxl+cQxam7xe0/ejYD063ZYCN9ya9TFEilblGr+8FVHh48iQieeePiotyhVuRCgD5zkC9Y0f1f4/WgvDQpZSbdGWLKtRSrs2omVCdnOh1RtXcFoyaZVIsUacYA4TFvrtGS0EPRf/GKKAklp+M2E1kiCSJ+FdpKVaqadvZ/9/lqzHkSR3gEnwkkWMj+iu7ThfGTBCsrdl4tuUSqDNnUnWCZswqGdqXgG1FpLllQkwYT5bh7UT8wJJA9QJNlMyAPq8JlriVd2FZt/b/qrtkbk7bgEnKTKwiGYT3KmLFRIvYWFpJSqI7MXVM4Gj2OmW2PFNYVK9ceROhCxRm1WqPDWFOPlg629MUpSk91AaI0kiDRgb/tJftAiBFzE34Ch3vjW/cc2nA8JU87VUb8sc4V8TYgLB1i8GcohZi2O6lijuAt6UpRASsp5IXxvjCNNHJe01hrRAKWDrlsCCLw9qizn4JAuUdjLcBofn5VpvzWinks/iV2n732DP0RKa1wlds0Hjcq0K2eMJk0o27q7jyG23bAUpAcMeUzCzTJxcHZkXO8kqxpO40cqYcr96XCKNiyzkHRfVA0m2XSzzCKsUdaTG5hiXn15w4BNLko7rRGuI1W6sT1UueOs1gghb7v7Nrk0tAm7aok1ogKcuSabsN9aFQ+B67zzpEvQGleJWGPWhs2nmuwGwy/z7zFZ+lcdMC/09y++1t6AyJpqPH0NU4hifkHpNewQ++PQAtrcrVGqYrO1Pkf9pMWuFlojGubPfh/1RZxEKN6Jb48ia54xa0SaSDnFIWqukLfcGqXMMc5uUPEGkxda8AbE5UNrXCVijVllLfUshU5t6dYoefM2YHPjnA00A3+Y/sOf/s5e/E4cDVOIoir/iPRx3lEx8K83iip8lzysRW7n0bs5ICQV6nMTelKZLSGFSE6wSY2h0hugvrzpzSYdYBHWKO+m8J64NGuU8gyVIQrnzDXU8ws3DvGvDJ5CxbWzh602xRIea0SGcfp293/+2t9On79t76vSiuE9FlXxlR1x0q3x+P6XyhONTVmwTNujJrfhYBfOXvBfpIbAEP7Ib3J0pt4HALoBrXGVtNkapak7bUlVdS0wETTvUBt23hPkoGEKUcpCn2JuIsBeUDfwr/RHofY/nGIgb1MPg2kKKtLFgwf+dBKsMVikEH7vyQ0gagyQXVGI9FnDKQ7KJh2gYtaLJtNA3rMGJVijtESk4ESEkMOdvbLT49wFp5jhz3q41jic4hcUhd2cA3uUOVyokLdH6WuJ6daI2qinWCXLM975nImbD5kpkqugNeaFfZm7jbBrhMccVEoMsUL1pAWflFk+HbLGrStydlVbV/J9q1zdybaiGfTGtMZKZJStu3ZU/w4FL3LKcDqcxFXRUaoECLVh1XCKoiunEJyvILWdPRsoAX/JVFkjtgaPdzCByZWNAb25XD1FrNDxDiYoRhvaR4U1zlcJ1iiZCczSHj39jw1UBOdRnKMUHgeF5mMDFXGtMboCCQrNMK/BJDI/c0i3RgecI7F/JyI030f+8zJvRLERjoi9qLI/xojQhtsYOmKNaD842ZlQpfKDRHTl+a2Qdwi5fFprjcGv15r8zMUa5ZJVqA3v7oc6SiABIn19eAoo65mhyy29sQ2UQDNrDObZ3AcUqZBIFmHcWKoqN2qVNd66/9h/IINJaP6nSAl4q0p/7C031xqrBhCYEvnTr2NOja0Rg3i/9y/CGlEThtNQaeP0YasbBaq6M66TdMQa15TWWqMM1QOzRoxwI1O6dGQX3jaMQ46+0t2GyUfJhN+jD2MEP0NRZyGhgTXKukVgIfHk2euRDCvSZ4XcsfieXodWWaN9CYMTa2evcllYqooJWYgIBV5+JtaI0kaJqcxzC3abD2yVeVs5/cHkwo3URxfWwxqjj11JXG81MyvSNtDGQGtcJa21xsjqyjytMXTZ6aV3IvcfZmE8qloB8/fO0IKt8fT52/48V631KSgNCexNYTjFybLhSrTHGiP2U7lYFzyKQJmjDgze/ujl0d1MdkMApO+/a3Q4RYWxgapYA2vc2UMxeufZSuQGK1ojWSpxa1zhwxvx1c65LKjKmm3o2kaza43m/h0bKMBKrBGdUWgQEO+q8gTv+O+Pr979tw1UImgqi1DUGmXR0pt/8yZe5BMO4ZVukrFFOfpgUml7iUhLLKffMWvsj+MDRxmtohDKtYXWSGYH1StdhQZp2mG2CQ2mYI2m38GPWYC4EjvcEEu41ihD1FAbbmaNCa91XYk1Bm84qjPckbffeRe7oh1ohTVmS4615C1AKGqN6JT9hWAGNKjqcbmxVHWsywElgzOOmowZOWxbRh4+S+iUNVY9gyE5QSH4yoHW2C5wmsf3vphdzsMGaAxoqLoJfzh3BqK1ZFsh5MFuSADtTR7GQiNJkKwvOU1id19+zLbmN6G+5rZWakb3ClojNKc7uYNtuOpTzBqmEEWV8NqOlVijPO/hPdK01VRF7mHxTj2j/hqzRnObIjrfWkIU7C5U/hFrlFttQ54KIcG4nPCqmtaIRoECR/5hVygW2wBDiUNdssaqN2kANAF/adMa24ZMqtAN4UzPou2R8wg5mvfxsxzbI+zFbjDI8Dbbat5AbzckgJ6oorHNosRkh1O0ebQ3m6dGyIGErXE+zzW+/5lMQZzEoaqPb0gYb1EkPJi8EmuUlxt4j7TOV7KDC5LDqVOB88Sscblvw/GvWM6oNOuCI2LOLXdgIQNo1DgXKJCsTPJ/O+qSNSbcZUprzNNqaww+Xl1Lpdop1pjVnlLtx7gyvzX0JIAXa40ad1VKaI0p+BsYNMPnqPIEXSrawwIJU269aZ8yXok1RnaaPrkXa/T24Gtkjd5szKIq6zrMvk3m2GGiumSNVRcaAa0xT9sXVIMddLo2zRrndMOOrAx7u5Lwo9a1cJPNVNWGJUw5Y7RGH+thjUgNmxqoP45YF3xFwqgpOntMVJesMeEr/7TGPK22RjlVPfMqEDSexhpMlmqN2CPyjBRWpa0r6VewIgQ79Hk8vxF79UbVJUMJU269G2yNEUNaC2s8ff42tjYQInof/MBRS5lr11HanRU2QTgQeGdoaYrWqKI1thA0AJgT2kBjIbpzXhdnjeDW/ccrF7oGm5sZkFIKdRn9MVqgDdcI9KEoWzdZKOF5Bg1WiGV+WUdrjA8C8uC0+ru8aAfadms0HfGM18XLSIF7ywpClUbmd/dRLMgSMozOAQWIyuOGhGiNKlrjhrBQa+wM0gOGJnY7exhz2HD1wVQpmHLCZz0kmK+Tba01vjy66x9kJFz+yUAv6U9kMGn4XONyrTE4GJr35wCldXtrl3FE1C7kv2zGaOae7NEaVbTGDYHWmIj0xd5eFRpO0cXYcDWRCYS3l8S+htPIe3AUGzIf0fzSWmu8cCP4SGJ6j+OfdUFRa2mPNUaey6zsstOBzcgjGeXjNS+qDU1PbSwnCkRrVNEaNwRaYyLBkT40nJ48G3v6MIQUvtdaoOgdJRkSstx6W2yNhRqVl3l9aKi/ziN9lrf3hKIfo2+PNWJq6y+ERgYQQoq6PGWs+nqXXMf15o3WqKI1bgi0xkSkE/Q2M1X9z38H262qN468DjRDQq6VNcobUEOjgX51hwWCq6lVfVZ7rFHsB2F8Jw6mYgPNjP/8Vn3ZA2UYKl5ao4jW2E4efv3s0vufzSLnYszirBE9UfrbcOav3f251+DgIxyq/hjzvMS5o/TvkZsGk7shCVxOpMXWCIJHbb68ES9AcVbk2ZtC1YW69lgjcoIT5M9JtOhq4f+yymASGXLJlW+ECRTvKq0xbRHFgdY4L9bAGqXqoCnifDdTb+x5Gw5+Pzrr87VGqeXeZrZoYad17nhMpGKeB+HU7O6j0GwEH+hMUYwV6bz0TqI5SeByCbfbGmNL04NJ5J4meWNfqEbt7FVWzvZYIwhOzqDeGMnacGmUF6KDBxu9nCk1M3RqVmuNUHS13AutcV6sx4JqrFFVqlQ7F26NGnfJWtjar9wAUln4JgB6RkzQ0QLh0Bino5X+7i/3Zd6JraHeR9Wr8TSI3ypmtsZyVxuigTWKnYTWVCHzbj9UvHweYIoyz0ZuvcaGH9Pu7fRPpKClW6OsqUYOx7xlCS3Ihg6A7vvW/cdSJ0vLsEFrDHwSEkmJL0bq9lKsEdn23wQE+W7UQvhIXUXrW4Q1Vt433j3WwxpjjapSG2KNvRpPAtTCjmorCx8BcI5QdOhr9G/8oX87IR1tj1KuqcAq0CnL432I4ksT/Qu2wpIR0sYpEbRG86HHSFyc2cMnz5CB4O0kZmSNDKCuensu9M7+XatwRNpH7+5Lf60FHim9/jjeW2W5DXa75l24CFN5S7CCEsChIbwYtjdXwymmv1oCoSm4lL8eZll6vLv7CINzgX4c+1IhzYsHD+C7UoYIiWAoyd7YJprDvig1n6xqMEGpIimcX+QNJxqlJ4WcBcbefQeFgR0MG8WoVatcVlrO2OpvI+Y5XY0bKhMQXGqGzLAJZaKSEvB9lBRVTs+OHJe3hM3IAwGQ4ZCzBq3RvAALcUN1u5OszW04Mk5El4SqXFf95S6oomKpPSxTvTEGrYurtWgVshctk/lqexRfq8FsEt2B9O/akUWuVkJaQxByd99rtzjXEsCJpTqKi31hj1nvY3tkpIkwpjoVYuWFrVpKu/voUp0plJRhZOKo0kPL/+uVMSFUNpt0DmTbFpcJFsstpIeMP6JXqdFjytQ/KwGN4pWWAHJuihFtB52pTcUgF/a8FpIJmzTbSEf3pf/VX7A1i7s9Klf42PISImKTqRs22fwm/axNlnheiJXl5+iLbKhdhWqJAN64EH7XrWYX3hMnnVs+P2VphiGk0zv+YAtKAFXaZgObjjLpFzZp4ezuI/Nlmw9aI6SJ419zFMiwjdNd1sYa0cYGb3/UTBduFF6HXcMazbARI1ZE8Qo9kVPR8ctKVLlaMiNyDUNb5lHJzEG9ceV8UeYoejt++q5NM/b29TFrVOle+mN9hTpOrqzRaZTEDGgwM9jXnWbIiD5XtRrKPDYTmoIc76JWcUXX/dCHWlNPLwENuT0qr5Pb9pWevZB89zPL+CNewlne8jIzKnEX53ef1BrRq1ibST8Q3fVwWnb04GfLvMp9y0zM7MyRoyfmRLPhuzUpZo0q3WTC2DjdZW2scY7A0o4rYpU1ivDfkBY5V2sbchkD7VB7hBllEslaeITCyUpXY2tUHX1dxFpjpLMIKXDtE1NJOZwGCUKIZW4JDvkiEGtMOUBH0VfGP4p8vDOuwCVkOac6BXHC15LvbYU4X5iW1aui/TFOClpx4pz+2BoblTNU7jHsXbKJKlqjTBmdACkaTptYY042TnehNSZYY0jmA+WbY40AbVjcZZbOHR2K+WBCpH/P0zFrBPZqUN00kW1zN4SzSuHQ0BqjzwnM3RqBLELMOMzq+e97kjkuUk4pBJwCUxVRpPAJuzgRkTlli7BGcDzdrxStcSnQGmmNtcH08dev/136Ne3iK9uShkFg0wUf1LlNX06WXl+spf44aI1ZniMqL6g6ASo1nEbumEUB2slNvAB1kwmGciv3ZWWOfbeWTPo2iRJ2QdWJkiLfgmoGOmK5NLi7L81NwzuHXxbC4FzgAPtjND3neeWMT758KhdHQ8nqj0jkzLXMXGVFJDvGLEwmrQCIMpwWFlTzYVJk0vEOClHTpHIiz7qvfIbzwqbcDXfWGrP00zWYOJeBgUydsXfNQJVsnO5Ca0ywxlK1sPJ9Tx9VPC/76xG1tibKRl46aFro2uSeQBQR3EtbtSP8CO3uIxjMptwaK0EvgEKGP9USbAkmYZPIgQxjkxO4LITRThPFC89wtiaq8q59DBG2Ln+MfUkBakGht82EIjUfiMCBpJfbq+YjEvlsJMppBXkwa2xwCiDEqrwEjs4dpxiFLF08KgyOOl8IEIoFP2LTGckD8qkVKV7z0SpRFAjvpom/zV2jSCRvURh26DFKNs5cy5chqi5OJYSThXOh7f3l0d1m5QxFBtMoCnt/crk+aDmYO7GzAQeSQngn/RSdOnfTa41OsIhsnO7STWtE1UHXlilbf1Dwi9Qztb1KazTdekho0vmKjuaK2UC2FRXI2TUGm/mtTu1EatnWFKFVeA1gyWAcjQO5ePAAmUH3gfJU4b8oahiADrQboyOABrLxczgB4moQxZGmEAfBMC1DAaKqoGfMhF9QbomJZGS7biCbhA8nZLps/ARgVDhkKF8IEIpFf29WGoiIyWWW2q37j1Ha3nQ0w2jOyAn+1f96Q4JsawPZJMJog8ryrNJyQObzHQ5wEk+XjV/ECRORjdBdummN6JdleURHXkt+G46JZVV61hA5Od66PUJDtRsM9hpJojBtrfMtXEIIIYl00xrlPkCYx5G9OUtbi7XG/C3gpQ/yWWs82uqxxixupQK3IRBCCJkRWuN6WqN56NtGI4QQMldojW2yRr38jgBxmextwnI/IYSsBFpjxW04J89eh715dfHgAQLnLWpGa3zl3cPf/eV+ij6pf6snIYSQRGiNFdYown+96rlvw5nRGgkhhLQBWmOCNYZUeuSf1kgIIR2A1khrJIQQUoDWWGWN5q03IdEaCSGke9Aao9YIXzxz/B3RsrAjWiMhhHQMWmPUGpf78AYhhJA2QGukNRJCCCnQXWvM2RutkRBCSDrdtMan//3f828ciMOZm2hojYQQQtLppjUqr974p1jjYEJrJIQQkk6XrRF88uXTX5iPkdr/G2iNhBBCInTcGgHsCm5n/2OgNRJCCInQfWssQ2skhBASYVOtsWfu0IE/rdAaS1/5J4QQ0gY20RoPnzw7ff62WBQ0mCzbGs2dQUj51LmbyIndQAghpDVsojUqn3z59MXRPRjV8q0Rezx49I39iRBCSMvYXGtUYGaOS83TGgcTxxpv3X/8xbff2/8QQghpJZtujWXEGreuiDsOJlATa9SlWqSwPXrz4KHdRgghZE2gNXrAxO7qp1+9euOfz79xANlfE4A1njx7/fT524gLi+WlREIIWUdojYQQQkgBWiMhhBBSgNZICCGEFKA1EkIIIQVojYQQQkgBWiMhhBBSgNZICCGEFKA1EkIIIQVojYQQQkgBWiMhhBBSgNZICCGEFKA1EkIIIQVojYQQQkgBWiMhhBBSgNZICCGEFKA1EkIIIQVojYQQQkgBWiMhhBBSgNZICCGEFKA1EkIIIQVojYQQQkiOn376P+pUwmFNtX9xAAAAAElFTkSuQmCC'
$banner_iconbytes = [Convert]::FromBase64String($banner_base64)
$banner_stream = [System.IO.MemoryStream]::new($banner_iconbytes, 0, $banner_iconbytes.Length)
$Form.BackgroundImage = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($banner_stream).GetHIcon()))
$Form.BackgroundImageLayout = "None"
$Form.StartPosition = "CenterScreen"
$Form.TopMost = $true
$tooltip = New-Object System.Windows.Forms.ToolTip
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "Location"
$Label1.AutoSize = $false
$Label1.Height = '50'
$Label1.Width = '180'
$Label1.Location = New-Object System.Drawing.Size(10,160)
$Label1.Font = $font1
$tooltip.SetToolTip($Label1, "Select at least one location to continue")
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.AutoCheck = $false
$checkBox1.Location = New-Object System.Drawing.Point(550,160)
$checkBox1.BackColor = [System.Drawing.Color]::Red
$checkBox1.Height = '20'
$checkBox1.Width = '20'
$form.Controls.Add($checkBox1)
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Location = New-Object System.Drawing.Point(550,250)
$checkBox2.AutoCheck = $false
$checkBox2.BackColor = [System.Drawing.Color]::Red
$checkBox2.Height = '20'
$checkBox2.Width = '20'
$form.Controls.Add($checkBox2)
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "Alert message"
$Label2.AutoSize = $false
$Label2.Height = '50'
$Label2.Width = '180'
$Label2.Font = $Font1
$Label2.Location = New-Object System.Drawing.Size(10,245)
$tooltip.SetToolTip($Label2, "Enter the alert message you would like to send")
$inputbox1 = New-Object System.Windows.Forms.TextBox
$inputbox1.BackColor = [System.Drawing.Color]::Beige
$inputbox1.Multiline = $true
$alert = $inputbox1.Text
$inputbox1.Location = New-Object System.Drawing.Size(180,250)
$inputbox1.Height = '180'
$inputbox1.Width = '350'
$inputbox1.Font = $Font2
$inputbox1.AcceptsReturn = $true
$inputbox1.MaxLength = '180'
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(180,450)
$okButton.Height = '40'
$okButton.Width = '350'
$okButton.Text = 'NOTIFY'
$okButton.Enabled = $False
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(180,160)
$listBox.Font = $font2
$listBox.Height = 80
$listBox.Width = 350
[void] $listBox.Items.Add('BGH-Campus')
[void] $listBox.Items.Add('Doctors Hospital')
[void] $listBox.Items.Add('Dublin Methodist')
[void] $listBox.Items.Add('Grady Memorial')
[void] $listBox.Items.Add('Grant Medical Center')
[void] $listBox.Items.Add('Groove city Medical Hospital')
[void] $listBox.Items.Add('Hardin Memorial Hospital')
[void] $listBox.Items.Add('Mansfield Hospital')
[void] $listBox.Items.Add('Mansfield Hospital MOB')
[void] $listBox.Items.Add('Marion General Hospital')
[void] $listBox.Items.Add('Marysville Hospital')
[void] $listBox.Items.Add('MedCentral Hospital')
[void] $listBox.Items.Add('Metro Data Center')
[void] $listBox.Items.Add('New Distro Center')
[void] $listBox.Items.Add('OAC1-Campus')
[void] $listBox.Items.Add('OBH1-Campus')
[void] $listBox.Items.Add('OhioHealth')
[void] $listBox.Items.Add('OhioHealth-VDI')
[void] $listBox.Items.Add('OPD1-Campus')
[void] $listBox.Items.Add('Perimeter Drive')
[void] $listBox.Items.Add('Pickerington Medical Campus(PMC)')
[void] $listBox.Items.Add('Riverside Methodist')
[void] $listBox.Items.Add('Shelby Hospital')
[void] $listBox.Items.Add('WAN Central Region')
[void] $listBox.Items.Add('WAN North Region')
[void] $listBox.Items.Add('WAN South Region')
[void] $listBox.Items.Add('WesterVille Hospital')
[void] $listBox.Items.Add('VanWert Hospital')
$listbox.SelectionMode = "MultiExtended"
$inputbox1.add_TextChanged({
   If($inputbox1.Text.Length -gt 0){
        $checkBox2.Checked = $true
        $checkBox2.BackColor = [System.Drawing.Color]::Green
   }
   Else{
        $checkBox2.Checked = $False
        $checkBox2.BackColor = [System.Drawing.Color]::Red
   }
})
# Event handler for list box
$listBox.add_SelectedIndexChanged({
    If($listBox.SelectedItems.Count -gt 0){
        $checkBox1.Checked = $true
        $checkBox1.BackColor = [System.Drawing.Color]::Green
   }
   Else{
        $checkBox1.Checked = $False
        $checkBox1.BackColor = [System.Drawing.Color]::Red
   }
})
$checkBox1.Add_CheckStateChanged({
    If ($checkBox1.Checked -and $checkBox2.Checked){
        $okButton.Enabled = $true
    }
    Else{
        $okButton.Enabled = $False
    }
})

$checkBox2.Add_CheckStateChanged({
    If ($checkBox1.Checked -and $checkBox2.Checked){
        $okButton.Enabled = $true
    }
    Else{
        $okButton.Enabled = $False
    }
})
$form.Controls.Add($listBox)
$form.Controls.Add($okButton)
$Form.Controls.Add($inputbox1)
$Form.MinimizeBox = $False
$Form.MaximizeBox = $False
$Form.Controls.Add($Label1)
$Form.Controls.Add($Label2)
$result = $form.ShowDialog()
$alertmsg = $inputbox1.Text
If(($result -eq [System.Windows.Forms.DialogResult]::OK)){    
        If ((($alertmsg -ne $null) -or (($alertmsg.Trim()).count -ne '0')) -and (($listbox.SelectedItems).count -gt '0')){
            Write-Output "Alert message received"
            #### ALERT MESSAGE INPUT RECEIVED, PUTTING THIS MESSAGE INTO SCCM
            If (Test-path "$env:SystemDrive\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"){
            Set-Location "$env:SystemDrive\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin"
	        Import-Module ".\ConfigurationManager.psd1"
            }
            Elseif (Test-Path "$env:SystemDrive\Program Files (x86)\ConfigMgrConsole\bin\ConfigurationManager.psd1"){
            Set-Location "$env:SystemDrive\Program Files (x86)\ConfigMgrConsole\bin"
	        Import-Module ".\ConfigurationManager.psd1"
            }
            Else { Write-Output ""}
	        New-PSDrive -Name "OHP" -PSProvider "CMSite" -Root "wapsccm01.ds.ohnet" -Description "PrimarySite" -EA SilentlyContinue
            $selected_locations = @($listbox.SelectedItems)
            $no_of_locations = $selected_locations.Count
            Write-Output "There are $no_of_locations selected and are listed below"
            Write-Output "#############################"
            Write-Output "$selected_locations"
            Write-Output "#############################"
            Write-Output "Fetching SCCM Script for sending alerts.."
            Set-Location OHP:
            $reqd_script = Get-CMScript -ScriptName "DONOTUSE-ISAlertTest" -Fast
            If ($reqd_script -ne $null){
                Write-Output "SendISAlert script located"
                $collection_ids = @()
                If ($no_of_locations -ne 0){
                    Write-Output "There are $no_of_locations selected"
                    foreach($location in $selected_locations){
                        Write-Output "Processing $location.."
                        switch ($location)
                        {
                            'BGH-Campus'{$collection_ids = $collection_ids + "OHP02A4A"}
                            'Doctors'{$collection_ids = $collection_ids + "OHP02A4C"}
                            'Dublin Methodist'{$collection_ids = $collection_ids + "OHP02A4D"}
                            'Grady Memorial'{$collection_ids = $collection_ids + "OHP02A66"}
                            'Grant Medical Center'{$collection_ids = $collection_ids + "OHP02A4E"}
                            'Groove city Medical Hospital'{$collection_ids = $collection_ids + "OHP02A4F"}
                            'Hardin Memorial Hospital'{$collection_ids = $collection_ids + "OHP02A50"}
                            'Mansfield Hospital'{$collection_ids = $collection_ids + "OHP02A51"}
                            'Mansfield Hospital MOB'{$collection_ids = $collection_ids + "OHP02A52"}
                            'Marion General'{$collection_ids = $collection_ids + "OHP02A53"}
                            'Metro Data Center'{$collection_ids = $collection_ids + "OHP02A54"}
                            'OAC1-Campus'{$collection_ids = $collection_ids + "OHP02A55"}
                            'OBH1-Campus'{$collection_ids = $collection_ids + "OHP02A56"}
                            'OhioHealth'{$collection_ids = $collection_ids + "OHP02A5D"}
                            'OhioHealth-VDI'{$collection_ids = $collection_ids + "OHP02A5C"}
                            'OPD1-Campus'{$collection_ids = $collection_ids + "OHP02A57"}
                            'Pickerington Medical Campus(PMC)'{$collection_ids = $collection_ids + "OHP02A5E"}
                            'Riverside Methodist'{$collection_ids = $collection_ids + "OHP02A5F"}
                            'Shelby Hospital'{$collection_ids = $collection_ids + "OHP02A62"}
                            'WAN Central Region'{$collection_ids = $collection_ids + "OHP02A63"}
                            'WAN North Region'{$collection_ids = $collection_ids + "OHP02A64"}
                            'WAN South Region'{$collection_ids = $collection_ids + "OHP02A65"}
                            'Vanwert'{$collection_ids = $collection_ids + "OHP02C0F"}
                            Default {}
                        }
                    }
                If ($collection_ids.count -ne '0'){
                    Write-Output "$collection_ids have been selected"
                    $alert = "$alertmsg" + "GHGH$no_of_locations impacted site listed belowGH"
                    foreach ($s in $selected_locations){
                        $alert = $alert + "GH$s"
                    }
                    $scriptParameters = @{
                            "isalert" = "$alert"
                    }
                    foreach($collection in $collection_ids){
                        Write-Output "Processing $collection.."
                        Invoke-CMScript -ScriptGuid "E42505F0-4F1A-4092-9D49-345405041129" -CollectionId "$collection" -ScriptParameter $scriptParameters
                    }
                }
            }
        }
        Else{

        }
    }
}
