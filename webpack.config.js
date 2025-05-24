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
      // your existing two HtmlWebpackPlugin instances…
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        publicPath: urlProd,
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        publicPath: urlProd,
      }),

      // copy assets + CSS + manifest + extra HTML
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              return content
                .toString()
                .replace(/https:\/\/localhost:3006\//g, urlProd);
            },
          },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },

          // ← NEW: copy your static support pages unchanged
          { from: "src/taskpane/privacy.html", to: "privacy.html" },
          { from: "src/taskpane/support.html", to: "support.html" },
          { from: "src/taskpane/index.html", to: "index.html" },
        ],
      }),
    ],
  };
};