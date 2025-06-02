const path              = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// must match your GH-Pages URL for docs/
const urlProd = "https://zojad.github.io/predloga-sz2/";

module.exports = (env, options) => {
  const dev = options.mode === "development";
  return {
    mode: dev ? "development" : "production",
    devtool: dev ? "inline-source-map" : "source-map",

    entry: {
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
    },

    output: {
      path: path.resolve(__dirname, "docs"),
      filename: "[name].js",
      publicPath: urlProd,
      clean: true,
    },

    resolve: { extensions: [".js"] },

    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
      ],
    },

    plugins: [
      // generate your taskpane.html
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        publicPath: urlProd,
      }),

      // generate your commands.html
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        publicPath: urlProd,
      }),

      // copy over assets, manifest, css, and all static pages
      new CopyWebpackPlugin({
        patterns: [
          // your icon assets
          { from: "assets", to: "assets" },

          // manifest (replace localhost with prod URL)
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              return content
                .toString()
                .replace(/https:\/\/localhost:3006\//g, urlProd);
            },
          },

          // your taskpane stylesheet
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },

          // English static pages
          { from: "src/taskpane/index.html",            to: "index.html" },
          { from: "src/taskpane/privacypolicy.html",   to: "privacy.html" },
          { from: "src/taskpane/support.html",          to: "support.html" },
          { from: "src/taskpane/license.html",          to: "license.html" },

          // Slovene static pages
          { from: "src/taskpane/index-si.html",         to: "index-si.html" },
          { from: "src/taskpane/privacypolicy-si.html",to: "privacy-si.html" },
          { from: "src/taskpane/support-si.html",       to: "support-si.html" },
          { from: "src/taskpane/license-si.html",       to: "license-si.html" },
        ],
      }),
    ],
  };
};
