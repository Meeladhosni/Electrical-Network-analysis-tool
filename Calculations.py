import numpy as np

class Y_bus_class:
    def __init__(self, num_buses):
        self.num_buses = num_buses
        self.Y_bus_matrix = np.zeros((self.num_buses, self.num_buses), dtype=np.complex128)

    def add_line_admittance(self, i, j, admittance, B, T):
        i_adj, j_adj = i - 1, j - 1
        B_half = 0.5 * B * 1j  # Convert to complex admittance for jB/2

        if T > 0:  # Transformer case

            # Add line admittance to off-diagonal elements
            self.Y_bus_matrix[i_adj][j_adj] -= admittance / T
            self.Y_bus_matrix[j_adj][i_adj] = self.Y_bus_matrix[i_adj][j_adj]
            # Adjust diagonal elements
            self.Y_bus_matrix[i_adj][i_adj] += admittance / T + ((1 / T) * (1 / T - 1) * admittance) + B_half
            self.Y_bus_matrix[j_adj][j_adj] += admittance / T + (1 - 1 / T) * admittance + B_half
        else:  # Line case

            # Add line admittance to off-diagonal elements
            self.Y_bus_matrix[i_adj, j_adj] -= admittance
            self.Y_bus_matrix[j_adj, i_adj] = self.Y_bus_matrix[i_adj, j_adj]
            # Adjust diagonal elements
            self.Y_bus_matrix[i_adj, i_adj] += admittance
            self.Y_bus_matrix[j_adj, j_adj] += admittance
            if B > 0:
                self.Y_bus_matrix[i_adj, i_adj] += B_half
                self.Y_bus_matrix[j_adj, j_adj] += B_half

    def add_shunt_element_admittance(self, i, Gs, Bs):
        bus_idx_adj = i - 1
        self.Y_bus_matrix[bus_idx_adj, bus_idx_adj] += Gs + 1j * Bs

    def get_Y_bus_matrix(self):
        return self.Y_bus_matrix


class Calc_class:
    @staticmethod
    def Power_generated(num_buses, bus_type_array, slack_bus_num, Y_bus_matrix, V_array,
                        PG_array, QG_array):
        slack_bus_index = slack_bus_num - 1
        # Calculate node currents
        I_node = np.dot(Y_bus_matrix, V_array)

        # Calculate complex power at each node
        S_node = V_array * np.conj(I_node)
        print("V_array", np.abs(V_array))
        print(np.angle(V_array, deg=True))

        print("S_node", S_node)

        # Initialize arrays for real and reactive power generation
        Pgen = np.zeros(num_buses)
        Qgen = np.zeros(num_buses)

        for k in range(num_buses):
            # Slack Bus
            if k == slack_bus_index:
                Pgen[k] = np.real(S_node[k])
                Qgen[k] = np.imag(S_node[k])
            # PV Bus
            elif bus_type_array[k] == 'PV':
                # For PV buses, only reactive power is adjusted based on power flow calculation
                Pgen[k] = PG_array[k]  # Use specified P value
                Qgen[k] = np.imag(S_node[k])  # Use calculated Q value
            # PQ Bus
            elif bus_type_array[k] == 'PQ':
                Pgen[k] = PG_array[k]  # Use specified P value
                Qgen[k] = QG_array[k]  # Use specified Q value

        return Pgen, Qgen

    @staticmethod
    def line_currents(num_buses, Y_bus_matrix, V_array):
        I_ij_array = np.zeros((num_buses, num_buses), dtype=np.complex128)  # current line array
        for i in range(num_buses):
            for j in range(num_buses):
                if i != j:
                    I_ij_array[i][j] = Y_bus_matrix[i][j] * (V_array[j] - V_array[i])

        return I_ij_array

    @staticmethod
    def line_power_flows(num_buses, V_array, I_ij_array):
        S_ij_array = np.zeros((num_buses, num_buses), dtype=np.complex128)  # Power flow line array

        for i in range(num_buses):
            for j in range(num_buses):
                if i != j:
                    S_ij_array[i][j] = V_array[i] * np.conj(I_ij_array[i][j])  # Power flow from i to j

        return S_ij_array

    @staticmethod
    def line_losses(PG_array, QG_array, PL_array, QL_array):

        total_P_loss = np.sum(PG_array) - np.sum(PL_array)
        total_Q_loss = np.sum(QG_array) - np.sum(QL_array)

        return total_P_loss, total_Q_loss

    @staticmethod
    def check_power_balance(bus_type_array, PG_array, PL_array, total_P_loss):
        """ Check if the total generation equals total load plus losses.

        Returns:
        bool: True if power balance is maintained, False otherwise.
        """

        total_P_generation = sum(PG_array)
        total_P_load = sum(PL_array[bus_type_array == 'PQ'])
        # total_losses = self.calculate_total_losses().real
        result_power_balance = np.isclose(total_P_generation, np.abs(total_P_load) + np.abs(total_P_loss), atol=1e-2)
        return result_power_balance, total_P_generation, total_P_load + np.abs(total_P_loss)

    @staticmethod
    def line_fault_power_flows(num_buses, V_array, I_ij_array):
        S_ij_array = np.zeros((num_buses, num_buses), dtype=np.complex128)  # Power flow line array

        for i in range(num_buses):
            for j in range(num_buses):
                if i != j:
                    S_ij_array[i][j] = V_array[i] * np.conj(I_ij_array[i][j])  # Power flow from i to j

        return S_ij_array

    @staticmethod
    def fault_line_currents(num_buses, Y_fault, V_fault_array):
        print("V_fault_array", V_fault_array)
        print("Y_fault", Y_fault)

        # Create a matrix where each row is the voltage array
        V_matrix = np.tile(V_fault_array, (num_buses, 1))
        print("V_matrix", V_matrix)

        # Calculate the voltage difference between each pair of buses
        # V_matrix.T (transpose) makes columns into rows and vice versa
        # The subtraction of V_matrix from its transpose gives the voltage difference matrix
        I_ij_array = Y_fault * (V_matrix.T - V_matrix)

        # Set the diagonal elements to 0 to avoid calculating self-currents
        np.fill_diagonal(I_ij_array, 0)
        print("I_ij_array", np.abs(I_ij_array))
        return I_ij_array
