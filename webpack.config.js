const path              = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// this must match your GitHub‐Pages URL for the /docs folder
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
      // ← this is the key change
      publicPath: urlProd,
      clean: true,
    },

    resolve: {
      extensions: [".js"],
    },

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
      // Taskpane page
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        // ensure the injected <script> tags use absolute paths
        publicPath: urlProd,
      }),

      // Commands (hidden iframe)
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        publicPath: urlProd,
      }),

      // copy over your static assets and rewrite localhost → GH-Pages
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              return content
                .toString()
                // replace your dev host with the prod root
                .replace(/https:\/\/localhost:3006\//g, urlProd);
            },
          },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
        ],
      }),
    ],
  };
};

