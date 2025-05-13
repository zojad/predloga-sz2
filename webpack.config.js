
const path              = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlProd = "https://zojad.github.io/predloga-sz2/";

module.exports = (env, options) => {
  const dev = options.mode === "development";

  return {
    devtool: "source-map",

    entry: {
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
    },

    output: {
      path: path.resolve(__dirname, "docs"),
      filename: "[name].js",
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
      // <-- taskpane.html only gets the taskpane bundle
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],        //<— only “taskpane”
      }),

      // <-- commands.html only gets the commands bundle
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],        //<— only “commands”
      }),

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
        ],
      }),
    ],
  };
};
