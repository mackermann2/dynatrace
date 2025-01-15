# -*- coding: utf-8 -*
# Supprime le délai d'animation de la souris
Settings.MoveMouseDelay = 0  
# Réduit le délai de saisie à 0 pour une saisie rapide
Settings.TypeDelay = 0  

def stopHandler(event):
    exit() # we want to exit the running script

Env.addHotkey(Key.F1,KeyModifier.ALT+KeyModifier.CTRL, stopHandler)

while True:
    try:
        if exists("1736892375320.png", 5):                        
            if exists(Pattern("1736893134679.png").similar(0.76), 5):
                type("matthieu.ackermann@externe.domaine.com")
                #wait(1)  # Petite pause avant la prochaine étape

            # Cliquer sur le bouton "Next"
            if exists("1736892407132.png", 5):
                click("1736892407132.png")

            # Saisir le mot de passe
            if exists("1736892453368.png", 5):
                send("TOUTOUYOUTOU")
                wait(1)
               
            if exists("1736892478687.png", 5):
               click("1736892478687.png")
        # Pause avant de refaire une boucle
        #wait(2)
    except Exception as e:
        # Gestion des erreurs éventuelles
        print("Erreur détectée : "+str(e))
        wait(5)
