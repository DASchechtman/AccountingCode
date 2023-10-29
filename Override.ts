interface String {
    toNumber(): number
}

String.prototype.toNumber = function () {
    return Number(this)
}