# server = 'https://ts7.x1.international.travian.com'
# my_login = "alexeenkoag1987@gmail.com"
# my_password = "Kontik-97"
server = 'https://ts4.x1.arabics.travian.com'
my_login = "alexeenkogi63@gmail.com"
my_password = "Papik-1963"
race = ['Romans', 'Huns', 'Teutonic'][0]  # поставить индекс рассы

raid = server + '/build.php?gid=16&tt=2&eventType=4'
normal_attack = server + '/build.php?gid=16&tt=2&eventType=5'
karte = server + '/karte.php'

if race == 'Romans':
    troops = ['', 'Legionnaire','Praetorian','Imperian','Equites Legati', 'Equites Imperatoris','Equites Caesaris',
              'empty','empty','empty','empty','Hero']
elif race == 'Teutonic':
    troops = ['', 'Maceman', 'Spearman', 'Axeman', 'Scout', 'Paladin', 'Teutonic Knight',
              'empty', 'empty', 'empty', 'empty', 'Hero']
else:
    troops = ['', 'Mercenary', 'Bowman', 'Spotter', 'Steppe Rider', 'Marksman','Marauder','empty','empty','empty',
              'empty','Hero']