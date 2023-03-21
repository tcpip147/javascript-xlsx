import { XMLParser, XMLBuilder } from "fast-xml-parser";
import _ from "lodash";
import DefaultXlsx from "./DefaultXlsx";

export default class Xlsx {

    constructor(document) {
        this.document = document;
    }

    afterNodeKey(path, key) {
        path = path.split("|");
        let node = this.document;
        let parent = {
            key: path[0],
            value: node
        };
        const newObj = {};
        _.forEach(path, (p, i) => {
            if (i == path.length - 1) {
                for (let k in node) {
                    newObj[k] = node[k];
                    if (k == path[i]) {
                        if (node[key] == null) {
                            newObj[key] = {};
                        }
                    }
                }
                parent.value[parent.key] = newObj;
            } else {
                parent = {
                    key: p,
                    value: node
                };
                node = node[p];
            }
        });
    }

    getNode(path) {
        path = path.split("|");
        let node = this.document;
        _.forEach(path, p => {
            if (node == null) {
                return null;
            }
            node = node[p];
        });
        if (Array.isArray(node)) {
            console.warn("The getNode function is not guaranteed to return an array when the return value is not array. It is recommended to use the getNodes function.");
        } else {
            return node;
        }
    }

    getNodes(path) {
        path = path.split("|");
        let node = this.document;
        _.forEach(path, p => {
            if (node == null) {
                return null;
            }
            node = node[p];
        });
        if (Array.isArray(node)) {
            return node;
        } else {
            const res = [];
            if (node != null) {
                res.push(node);
            }
            return res;
        }
    }

    setNode(path, newNode, isArray) {
        path = path.split("|");
        let node = this.document;
        _.forEach(path, (p, i) => {
            if (i == path.length - 1) {
                node[p] = newNode;
            } else {
                if (node[p] == null || node[p] === "") {
                    if (isArray && i == path.length - 2) {
                        node[p] = [];
                    } else {
                        node[p] = {};
                    }
                }
                node = node[p];
            }
        });
    }

    appendNode(path, newNode, isArray) {
        path = path.split("|");
        let node = this.document;
        _.forEach(path, (p, i) => {
            if (i == path.length - 1) {
                if (isArray && node[p] == null) {
                    node = [];
                }
                if (Array.isArray(node[p])) {
                    node[p].push(newNode);
                } else {
                    if (node[p] == null) {
                        node[p] = newNode;
                    } else {
                        const temp = node[p];
                        node[p] = [];
                        node[p].push(temp);
                        node[p].push(newNode);
                    }
                }
            } else {
                if (node[p] == null || node[p] === "") {
                    if (isArray && i == path.length - 2) {
                        node[p] = [];
                    } else {
                        node[p] = {};
                    }
                }
                node = node[p];
            }
        });
    }

    removeNode(path, condition) {
        path = path.split("|");
        let node = this.document;
        _.forEach(path, (p, i) => {
            if (i == path.length - 1) {
                if (condition != null) {
                    const conditionCount = Object.keys(condition).length;
                    if (Array.isArray(node[p])) {
                        const removing = [];
                        _.forEach(node[p], (n, j) => {
                            let removable = 0;
                            for (let k in condition) {
                                if (condition[k] == n[k]) {
                                    removable++;
                                }
                            }
                            if (removable == conditionCount) {
                                removing.push(j);
                            }
                        });
                        _.forEach(removing, j => {
                            node[p].splice(j, 1);
                        });
                    } else {
                        let removable = 0;
                        for (let k in condition) {
                            if (condition[k] == node[p][k]) {
                                removable++;
                            }
                        }
                        if (removable == conditionCount) {
                            delete node[p];
                        }
                    }
                } else {
                    console.log(node);
                    delete node[p];
                }
            } else {
                node = node[p];
            }
        });
    }
}