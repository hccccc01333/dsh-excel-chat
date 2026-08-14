/**
 * dsh tool-DSL schema for FormulaIR. Validates the model's IR payload before
 * `excel_compile_formula` runs.
 */
export declare const formulaIrSchema: {
    readonly oneOf: readonly [{
        readonly type: 'object';
        readonly additionalProperties: false;
        readonly properties: {
            readonly operation: {
                readonly type: 'string';
                readonly enum: readonly ['binary'];
                readonly required: true;
            };
            readonly left: {
                readonly oneOf: readonly [{
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['column'];
                            readonly required: true;
                        };
                        readonly column: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['cell'];
                            readonly required: true;
                        };
                        readonly cell: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['constant'];
                            readonly required: true;
                        };
                        readonly value: {
                            readonly type: 'number';
                            readonly required: true;
                        };
                    };
                }];
                readonly required: true;
            };
            readonly right: {
                readonly oneOf: readonly [{
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['column'];
                            readonly required: true;
                        };
                        readonly column: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['cell'];
                            readonly required: true;
                        };
                        readonly cell: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['constant'];
                            readonly required: true;
                        };
                        readonly value: {
                            readonly type: 'number';
                            readonly required: true;
                        };
                    };
                }];
                readonly required: true;
            };
            readonly operator: {
                readonly type: 'string';
                readonly enum: readonly ['+', '-', '*', '/'];
                readonly required: true;
            };
        };
    }, {
        readonly type: 'object';
        readonly additionalProperties: false;
        readonly properties: {
            readonly operation: {
                readonly type: 'string';
                readonly enum: readonly ['aggregate'];
                readonly required: true;
            };
            readonly metric: {
                readonly type: 'string';
                readonly required: true;
            };
            readonly function: {
                readonly type: 'string';
                readonly enum: readonly ['SUMIFS', 'AVERAGEIFS', 'COUNTIFS', 'SUM'];
                readonly required: true;
            };
            readonly filters: {
                readonly type: 'array';
                readonly required: true;
                readonly items: {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly column: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                        readonly value_from: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                };
            };
        };
    }, {
        readonly type: 'object';
        readonly additionalProperties: false;
        readonly properties: {
            readonly operation: {
                readonly type: 'string';
                readonly enum: readonly ['ratio'];
                readonly required: true;
            };
            readonly numerator: {
                readonly oneOf: readonly [{
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['column'];
                            readonly required: true;
                        };
                        readonly column: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['cell'];
                            readonly required: true;
                        };
                        readonly cell: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['constant'];
                            readonly required: true;
                        };
                        readonly value: {
                            readonly type: 'number';
                            readonly required: true;
                        };
                    };
                }];
                readonly required: true;
            };
            readonly denominator: {
                readonly oneOf: readonly [{
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['column'];
                            readonly required: true;
                        };
                        readonly column: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['cell'];
                            readonly required: true;
                        };
                        readonly cell: {
                            readonly type: 'string';
                            readonly required: true;
                        };
                    };
                }, {
                    readonly type: 'object';
                    readonly additionalProperties: false;
                    readonly properties: {
                        readonly kind: {
                            readonly type: 'string';
                            readonly enum: readonly ['constant'];
                            readonly required: true;
                        };
                        readonly value: {
                            readonly type: 'number';
                            readonly required: true;
                        };
                    };
                }];
                readonly required: true;
            };
        };
    }];
};
//# sourceMappingURL=ir-schema.d.ts.map