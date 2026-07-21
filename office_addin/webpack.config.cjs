const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (_env, argv) => {
  const development = argv.mode === "development";
  const devServer = development
    ? {
        port: 3000,
        hot: false,
        server: {
          type: "https",
          options: await devCerts.getHttpsServerOptions(),
        },
        headers: {
          "Access-Control-Allow-Origin": "*",
        },
        static: false,
      }
    : undefined;

  return {
    entry: "./src/taskpane.ts",
    output: {
      filename: "taskpane.js",
      path: path.resolve(__dirname, "dist"),
      clean: true,
    },
    devtool: development ? "source-map" : false,
    devServer,
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
            options: {
              onlyCompileBundledFiles: true,
            },
          },
        },
      ],
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          { from: "public/taskpane.html", to: "taskpane.html" },
          { from: "public/styles.css", to: "styles.css" },
          { from: "public/config.json", to: "config.json" },
          {
            from: "public/assets",
            to: "assets",
            globOptions: { ignore: ["**/*.b64"] },
          },
          {
            from: "public/assets/icon.png.b64",
            to: "assets/icon.png",
            transform(content) {
              return Buffer.from(content.toString("utf8").trim(), "base64");
            },
          },
          { from: "manifest.xml", to: "manifest.xml" },
        ],
      }),
    ],
  };
};
