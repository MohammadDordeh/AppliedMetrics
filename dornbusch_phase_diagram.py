import numpy as np
import matplotlib.pyplot as plt


def compute_phase_variables(delta=0.5, theta=0.3, lambda_=0.8, grid_range=(-2, 2), points=25):
    """Compute phase diagram variables for the Dornbusch overshooting model."""
    denominator = delta * lambda_ + theta * lambda_ - theta
    a1 = (delta + theta) / denominator
    a2 = delta / denominator

    e_vals = np.linspace(grid_range[0], grid_range[1], points)
    p_vals = np.linspace(grid_range[0], grid_range[1], points)
    E, P = np.meshgrid(e_vals, p_vals)

    dE = a1 * E - a2 * P
    dP = delta * (E + P)

    norm = np.sqrt(dE**2 + dP**2)
    dE_norm = dE / norm
    dP_norm = dP / norm

    return {
        "E": E,
        "P": P,
        "dE_norm": dE_norm,
        "dP_norm": dP_norm,
        "e_vals": e_vals,
        "a1": a1,
        "a2": a2,
    }


def plot_phase_diagram(vars_dict):
    """Plot the phase diagram using the computed variables."""
    E, P = vars_dict["E"], vars_dict["P"]
    dE_norm, dP_norm = vars_dict["dE_norm"], vars_dict["dP_norm"]
    e_vals = vars_dict["e_vals"]
    a1, a2 = vars_dict["a1"], vars_dict["a2"]

    fig, ax = plt.subplots(figsize=(10, 7))
    ax.quiver(E, P, dE_norm, dP_norm, angles="xy", scale_units="xy", scale=1.5, color="steelblue")
    ax.axhline(0, color="black", lw=1)
    ax.axvline(0, color="black", lw=1)

    ax.plot(e_vals, (a1 / a2) * e_vals, "r--", label="dE/dt = 0")
    ax.plot(e_vals, -e_vals, "g--", label="dP/dt = 0")

    ax.set(
        title="Phase Diagram of Dornbusch Overshooting Model",
        xlabel="Exchange Rate Deviation (e - ē)",
        ylabel="Price Level Deviation (p - p̄)",
        xlim=(e_vals[0], e_vals[-1]),
        ylim=(e_vals[0], e_vals[-1]),
    )
    ax.legend()
    ax.grid(True)
    return fig, ax


if __name__ == "__main__":
    vars_dict = compute_phase_variables()
    plot_phase_diagram(vars_dict)
    plt.show()
