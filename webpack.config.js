/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const isProd = process.env.NODE_ENV === "production";

module.exports = {
  entry: {
    taskpane: path.resolve(__dirname, "src/taskpane/index.tsx"),
    commands: path.resolve(__dirname, "src/commands/commands.ts"),
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js", ".jsx"],
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: path.resolve(__dirname, "src/taskpane/taskpane.html"),
      chunks: ["taskpane"],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: path.resolve(__dirname, "src/commands/commands.html"),
      chunks: ["commands"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: path.resolve(__dirname, "assets"),
          to: path.resolve(__dirname, "dist/assets"),
          noErrorOnMissing: true,
        },
      ],
    }),
  ],
  devtool: isProd ? false : "source-map",
  devServer: {
    port: 3000,
    server: "https",
    hot: true,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
  mode: isProd ? "production" : "development",
};
