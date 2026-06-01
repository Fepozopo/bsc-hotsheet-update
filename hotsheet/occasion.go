package hotsheet

import "strings"

// Token lists used by mapOccasion.
var (
	everTokens = []string{
		"ALL OCCASION",
		"BABY",
		"BAPTISM-COMMUNION",
		"BIRTHDAY",
		"BLANK",
		"CAMP",
		"CANCER",
		"CONGRATULATIONS",
		"ENCOURAGEMENT",
		"FRIENDSHIP",
		"GET WELL",
		"KID BIRTHDAY",
		"LOVE",
		"MENOPAUSE",
		"MISS YOU",
		"NEW HOME",
		"PET SYMPATHY",
		"SORRY",
		"SYMPATHY",
		"TEACHER APPRECIATION",
		"THANK YOU",
		"THINKING OF YOU",
		"WEDDING ANNIVERSARY",
	}
	winterTokens = []string{
		"CHRISTMAS",
		"HALLOWEEN",
		"HANUKKAH",
		"HOLIDAY",
		"THANKSGIVING",
		"VETERAN'S DAY",
		"VETERANS DAY",
	}
	springTokens = []string{
		"EASTER",
		"FATHER'S DAY",
		"FATHERS DAY",
		"GRADUATION",
		"INDEPENDENCE DAY",
		"MOTHER'S DAY",
		"MOTHERS DAY",
		"ST PATRICKS DAY",
		"ST. PATRICK'S DAY",
		"VALENTINE'S DAY",
		"VALENTINES DAY",
	}
)

// mapOccasion maps raw occasion text to one of: "Everyday", "Winter", or "Spring".
// The token lists are checked in season order so more specific holiday matches win before
// the broad Everyday fallback can claim the row.
func mapOccasion(occ string) string {
	o := strings.ToUpper(strings.TrimSpace(occ))
	if o == "" {
		return "Everyday"
	}
	for _, t := range springTokens {
		if strings.Contains(o, t) {
			return "Spring"
		}
	}
	for _, t := range winterTokens {
		if strings.Contains(o, t) {
			return "Winter"
		}
	}
	for _, t := range everTokens {
		if strings.Contains(o, t) {
			return "Everyday"
		}
	}
	return "Everyday"
}
