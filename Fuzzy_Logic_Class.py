class FuzzyLogic:
    def __init__(self, word, spell):
        self.word = word
        self.spell = spell
    
    def percentage(self):
        from fuzzywuzzy import fuzz
        from fuzzywuzzy import process

        return fuzz.ratio(self.word, self.spell)
