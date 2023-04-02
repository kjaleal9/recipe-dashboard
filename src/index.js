import React from "react";
import ReactDOM from "react-dom/client";
import { createBrowserRouter, RouterProvider } from "react-router-dom";
import "./index.css";
import App from "./App";
import RecipeSearch from "./Routes/RecipeSearch";
import RecipeEdit from "./Routes/RecipeEdit";
import Materials from "./Routes/Materials";

const router = createBrowserRouter([
  {
    path: "/",
    element: <App />,
    children: [
      {
        path: "recipes",
        element: <RecipeSearch />,
      },
      {
        path: "recipes/:RID/:Version/edit",
        element: <RecipeEdit />,
      },
      {
        path: "materials",
        element: <Materials />,
      },
    ],
  },
]);

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(
  <React.StrictMode>
    <RouterProvider router={router} />
  </React.StrictMode>
);
