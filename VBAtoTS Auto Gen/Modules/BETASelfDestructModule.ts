// TODO: Option Explicit ... Warning!!! not translated


KillMe() {
    // With...
    /* Warning! Labeled Statements are not Implemented */xlReadOnly;
    Kill.FullName.Close;
    false;
}

CheckDate() {
    let BetaDate: Date;
    BetaDate = HiddenSht.Range("BetaDate").Value;
    if ((BetaDate < Now)) {
        KillMe;
    }

}
