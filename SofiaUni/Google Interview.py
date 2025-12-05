import random
from typing import List, Tuple

class Deck:
    Suits = ["Hearts", "Spades", "Clubs", "Diamonds"]
    Ranks = ["A", "2", "3", "4", "5", "6", "7", "8", "9", "10", "J", "Q", "K"]

    def __init__(self):
        # build 52-card deck like "HeartsA", "Spades10"
        self.cards: List[str] = [suit + rank for suit in self.Suits for rank in self.Ranks]

    def draw_after_shuffle(self, n: int = 3):
        if n > len(self.cards):
            print("Not enough cards to draw")
            return [], self.cards

        random.shuffle(self.cards)     # shuffle in place
        card_draw = self.cards[:n]     # pick top n cards
        remain_card = self.cards[n:]   # the rest
        self.cards = remain_card       # update deck state

        print(f"I draw {n} card(s) from deck, and it remains {len(remain_card)} cards in deck")
        return card_draw, remain_card


# Example usage
deck = Deck()
hand, rest = deck.draw_after_shuffle(5)
print("Drawn:", hand)
print("Remaining:", len(rest))
