export default class IndexedLinkedList {

    #list;
    #first;
    #last;
    #size;
    #keyGen;

    constructor(option) {
        this.#list = {};
        this.#size = 0;
        this.#keyGen = option && option.keyGen;
    }

    addWithKeyGen(value, override) {
        if (this.#keyGen) {
            this.add(this.#keyGen(value), value, override);
        } else {
            this.add(this.#size, value, override);
        }
    }

    add(key, value, override) {
        if (this.#list[key] != undefined) {
            if (override) {
                this.#list[key].value = value;
                return;
            } else {
                throw "Key does already exists";
            }
        }

        if (this.#size == 0) {
            const item = {
                key: key,
                prev: undefined,
                value: value,
                next: undefined
            };
            this.#first = item;
            this.#last = item;
            this.#list[key] = item;
            this.#size++;
        } else {
            this.after(this.#last.key, key, value);
        }
    }

    after(at, key, value) {
        if (this.#list[at] == undefined) {
            throw "A target does not exists."
        }
        const item = {
            key: key,
            prev: undefined,
            value: value,
            next: undefined
        };
        this.#list[key] = item;
        const next = this.#list[at].next;
        if (next == undefined) {
            this.#last = item;
        } else {
            item.next = next;
            next.prev = item;
        }
        item.prev = this.#list[at];
        this.#list[at].next = item;
        this.#size++;
    }

    before(at, key, value) {
        if (this.#list[at] == undefined) {
            throw "A target does not exists."
        }
        const item = {
            key: key,
            prev: undefined,
            value: value,
            next: undefined
        };
        this.#list[key] = item;
        const prev = this.#list[at].prev;
        if (prev == undefined) {
            this.#first = item;
        } else {
            item.prev = prev;
            prev.next = item;
        }
        item.next = this.#list[at];
        this.#list[at].prev = item;
        this.#size++;
    }

    get(key) {
        return this.#list[key];
    }

    remove(key) {
        if (this.#list[key] != undefined) {
            const prev = this.#list[key].prev;
            const next = this.#list[key].next;
            if (prev != undefined) {
                prev.next = next;
            } else {
                this.#first = next;
            }
            if (next != undefined) {
                next.prev = prev;
            } else {
                this.#last = prev;
            }
            delete this.#list[key];
            this.#size--;
        }
    }

    first() {
        return this.#first;
    }

    last() {
        return this.#last;
    }

    each(callback) {
        let node = this.#first;
        for (let i = 0; i < this.#size; i++) {
            callback(node.key, node.value);
            node = node.next;
        }
    }

    eachRight(callback) {
        let node = this.#last;
        for (let i = 0; i < this.#size; i++) {
            callback(node.key, node.value);
            node = node.prev;
        }
    }

    toString() {
        const res = {};
        this.each((key, value) => {
            res[key] = value;
        });
        return res;
    }
}