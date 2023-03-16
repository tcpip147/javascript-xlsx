const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyPlugin = require("copy-webpack-plugin");
const TerserPlugin = require("terser-webpack-plugin");

module.exports = (env, argv) => {
    const config = {
        cache: false,
        entry: "./src/Entry.js",
        output: {
            filename: "javascript-xlsx.js",
            path: path.resolve(__dirname, "dist"),
            clean: true
        },
        module: {
            rules: [
                {
                    test: /\.?js$/,
                    enforce: "pre",
                    use: ["source-map-loader"]
                }
            ]
        },
        plugins: [
            new HtmlWebpackPlugin({
                template: path.join(__dirname, "./example/example.html"),
                filename: "example.html",
                inject: false
            }),
            new CopyPlugin({
                patterns: [{
                    from: "./src",
                    filter: async (resourcePath) => {
                        return /\.(css)$/i.exec(resourcePath);
                    }
                }],
            })
        ],
        devServer: {
            static: {
                directory: path.join(__dirname, 'dist'),
            },
            compress: false,
            port: 9999
        },
        optimization: {
            minimize: true,
            minimizer: [
                new TerserPlugin({ extractComments: false })
            ]
        }
    };

    if (env.ie11) {
        config.target = ["web", "es5"];
        config.output = {
            filename: "javascript-xlsx-ie11.js",
            path: path.resolve(__dirname, "dist-ie11"),
            clean: true
        };
        config.plugins[0] =
            new HtmlWebpackPlugin({
                template: path.join(__dirname, "./example/example-ie11.html"),
                filename: "example-ie11.html",
                inject: false
            });
        config.module.rules.push({
            test: /\.m?js$/,
            exclude: {
                and: [/node_modules/],
                not: [
                    /node_modules\\fast-xml-parser/,
                    /node_modules\\strnum/,
                ]
            },
            use: {
                loader: "babel-loader",
                options: {
                    presets: [
                        [
                            "@babel/preset-env", {
                                debug: false,
                                targets: {
                                    ie: "11"
                                },
                                modules: false,
                                corejs: "3",
                                useBuiltIns: "usage"
                            }
                        ]
                    ],
                    plugins: [
                        "@babel/plugin-transform-runtime",
                        "@babel/plugin-transform-modules-commonjs",
                    ]
                }
            }
        });
    }
    return config;
};