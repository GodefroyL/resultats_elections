

class departement:
    def __init__(self, id, name, code):
        self.id = id
        self.name = name
        self.code = code


class circonscription:
    def __init__(self, id, name, code, departement_id):
        self.id = id
        self.name = name
        self.code = code
        self.departement_id = departement_id
        self.list_candidates = []
        self.nombres_votants = None
        self.nombres_inscrits = None

    def get_candidates(self):
        # Placeholder for method to retrieve candidates for this circonscription
        return self.list_candidates
    
    def add_candidate(self, candidate, party=None, score_percent=None, score_number=None):
        self.list_candidates.append({
            "candidate": candidate,
            "party": party,
            "score_percent": score_percent,
            "score_number": score_number
        })
    
    def set_voting_numbers(self, votants, inscrits):
        self.nombres_votants = votants
        self.nombres_inscrits = inscrits
    
    def get_voting_numbers(self):
        return {
            "nombres_votants": self.nombres_votants,
            "nombres_inscrits": self.nombres_inscrits
        }
    
class candidate:
    def __init__(self, name, party=None, circonscription_id=None, score_number=None):
        self.name = name
        self.party = party
        self.circonscription_id = circonscription_id
        self.score_number = score_number
    
    def get_info(self):
        return {
            "name": self.name,
            "party": self.party,
            "circonscription_id": self.circonscription_id,
            "score_number": self.score_number
        }
    
    def get_score_percent(self):
        total_votes = self.circonscription_id.nombres_votants if self.circonscription_id else 0
        if total_votes > 0:
            return (self.score_number / total_votes) * 100
        return 0