export default class Utils {

    static alphabetToIndex(alphabet) {
        const match = alphabet.match(/([A-Z]+)/);
        let digit = 1;
        let index = 0;
        for (let i = match[1].length - 1; i >= 0; i--) {
            const column = match[1].charCodeAt(i) - 64;
            index += (column * digit);
            digit *= 26;
        }
        return index;
    };

    static indexToAlphabet(index) {
        let alphabet = "";
        let temp = index - 1;
        while (true) {
            if (temp < 26) {
                alphabet = String.fromCharCode(temp % 26 + 65) + alphabet;
                break;
            }
            alphabet = String.fromCharCode(temp % 26 + 65) + alphabet;
            temp = Math.floor(temp / 26) - 1;
        }
        return alphabet;
    };

    static regionToIndices(region) {
        const addresses = region.split(":");
        const start = Utils.addressToIndices(addresses[0]);
        const end = Utils.addressToIndices(addresses[1]);
        return _.concat(start, end);
    }

    static addressToIndices(address) {
        const match = address.match(/([A-Z]+)([0-9]+)/);
        return [Number(match[2] - 1), Number(Utils.alphabetToIndex(match[1]) - 1)];
    }
}