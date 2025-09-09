def checkInclusion(s1, s2):
    n = len(s1)
    m = len(s2)
    if n > m:
        return False

    s1_sorted = sorted(s1)

    for i in range(m - n + 1):
        window = sorted(s2[i:i + n])
        print(window)
        if window == s1_sorted:
            return True

    return False





s1 = "abb"
s2 = "ddc"
checkInclusion(s1, s2)