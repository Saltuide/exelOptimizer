import re

from spellchecker import SpellChecker


# noinspection PyMethodMayBeStatic
class AbbreviationFixer:

    def __init__(self, spell_checker: SpellChecker):
        self.spell_checker = spell_checker
        self.spell_checker.word_frequency.load_dictionary('Files/dict.json', encoding='utf-8')

    def correct_abbreviation(self, string: str) -> str:
        new_value = ''
        values = re.split('(\d+)', string)
        for value in values:
            if value.isnumeric():
                new_value += value
                continue

            if value == '':
                continue

            new_value += self.spell_checker.correction(value.strip())
        return new_value
