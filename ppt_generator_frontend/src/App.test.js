import { render, screen } from "@testing-library/react";
import App from "./App";

test("renders ppt generator header", () => {
  render(<App />);
  const heading = screen.getByText(/PPT Generator/i);
  expect(heading).toBeInTheDocument();
});
