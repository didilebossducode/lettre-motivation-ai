# Générateur de Lettre de Motivation

Cette application permet de réutiliser facilement un modèle de lettre de motivation en remplaçant uniquement certaines parties spécifiques.

## Comment ça marche

1. **Remplissez les informations** :
   - Entreprise
   - Poste
   - Durée
   - Date de début
   - Date du jour (pré-remplie avec aujourd'hui)
   - **Paragraphe personnalisé** : ce texte remplacera entièrement la partie surlignée en orange

2. **Dans votre modèle de lettre** :
   - Collez votre modèle de lettre dans la zone du milieu
   - Surlignez les parties à remplacer avec les bonnes couleurs :
     - 🟢 Vert : Nom de l'entreprise
     - 🔵 Bleu : Nom du poste
     - 🟣 Violet : Durée
     - 💛 Jaune : Date de début
     - 💗 Rose : Date du jour
     - 🟧 Orange : **Paragraphe entier** qui sera remplacé par votre texte personnalisé

3. **Génération de la lettre** :
   - L'application garde exactement le même modèle
   - Pour les parties en vert, bleu, violet, jaune et rose : elle remplace uniquement le mot ou groupe de mots surligné
   - Pour la partie en orange : elle remplace tout le paragraphe surligné par votre texte personnalisé
   - Le reste du texte reste exactement identique
   - L'ordre des paragraphes est préservé

## Exemple

Si votre modèle contient :
```
Je souhaite postuler au poste de [développeur Python] dans votre entreprise [Microsoft] pour une durée de [6 mois] à partir du [1er mars 2024].

[Je suis particulièrement intéressé par ce poste car il correspond parfaitement à mes compétences en développement. Mon expérience précédente chez XXX m'a permis de développer une expertise solide dans ce domaine.]

Paris, le [17/01/2024]
```

Et que vous remplissez :
- Poste : "ingénieur logiciel"
- Entreprise : "Google"
- Durée : "1 an"
- Date de début : "1er septembre 2024"
- Date du jour : "17/01/2024"
- Paragraphe personnalisé : "Ayant travaillé pendant 3 ans sur des projets similaires, je maîtrise parfaitement les technologies requises pour ce poste. Mon expérience en développement agile et en intégration continue serait un atout pour votre équipe."

La lettre générée sera :
```
Je souhaite postuler au poste de ingénieur logiciel dans votre entreprise Google pour une durée de 1 an à partir du 1er septembre 2024.

Ayant travaillé pendant 3 ans sur des projets similaires, je maîtrise parfaitement les technologies requises pour ce poste. Mon expérience en développement agile et en intégration continue serait un atout pour votre équipe.

Paris, le 17/01/2024
```

## Important
- Les parties en vert, bleu, violet, jaune et rose : seul le texte surligné est remplacé
- La partie en orange : tout le paragraphe surligné est remplacé par votre texte personnalisé
- L'ordre et la structure de la lettre sont préservés
- Le reste du texte reste exactement identique
